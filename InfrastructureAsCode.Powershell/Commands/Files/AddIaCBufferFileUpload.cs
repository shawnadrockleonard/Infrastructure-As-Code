using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    [Cmdlet(VerbsCommon.Add, "IaCBufferFileUpload")]
    [CmdletHelp("Uploads a large file via chunking to a library specified", Category = "Files")]
    public class AddIaCBufferFileUpload : IaCCmdlet
    {
        #region Parameters 

        /// <summary>
        /// The display name for the library
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind ListTitle { get; set; }

        /// <summary>
        /// The full path to the file
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string FileName { get; set; }

        /// <summary>
        /// The foldername in which the file will be uploaded
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string FolderName { get; set; }

        /// <summary>
        /// Should we overwrite the file if it exists
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public SwitchParameter Clobber { get; set; }

        #endregion


        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (!System.IO.File.Exists(this.FileName))
            {
                throw new System.IO.FileNotFoundException(String.Format("{0} {1}", this.FileName, InfrastructureAsCode.Core.Properties.Resources.FileDoesNotExist));
            }
        }

        /// <summary>
        /// Execute the command uploading the file to the root of the document library
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                LogVerbose("Entering Library Upload Cmdlet");
                var ctx = this.ClientContext;

                var w = ctx.Web;
                var l = ListTitle.GetList(ctx.Web);
                ctx.Load(w, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
                ctx.Load(l, listctx => listctx.RootFolder, listctx => listctx.RootFolder.ServerRelativeUrl);
                ClientContext.ExecuteQueryRetry();

                var webUri = new Uri(w.Url);
                var serverRelativeUrl = new Uri(webUri, Path.Combine(l.RootFolder.ServerRelativeUrl, FolderName));
                LogVerbose(string.Format("Uploading file to {0}", serverRelativeUrl.AbsoluteUri));


                if (UploadFileWithBuffer(ctx, l, serverRelativeUrl, FileName, 10))
                {
                    LogVerbose("Successfully uploaded {0}", FileName);
                }

            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in SetFileUpload File:{0}", this.FileName);
            }

        }

        private bool m_ensuredPath = false;
        private bool EnsureFolderPath(ClientContext m_clientContext, List library, string webListFolderUrl)
        {
            if (m_ensuredPath)
                return m_ensuredPath;

            // make sure the RootFolder object is loaded into memory
            library.EnsureProperties(lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl);

            var targetFolder = m_clientContext.Web.GetFolderByServerRelativeUrl(webListFolderUrl);
            try
            {
                m_clientContext.Load(targetFolder);
                m_clientContext.ExecuteQueryRetry();
                m_ensuredPath = true;
            }
            catch (Exception)
            {
                // Folder Not Found
                // Enumerate the Folder structure creating the desired destination 
                targetFolder = library.RootFolder;
                var relativeUri = TokenHelper.EnsureTrailingSlash(targetFolder.ServerRelativeUrl);
                var targetFolderPath = webListFolderUrl.Replace(relativeUri, string.Empty).Split(new char[] { '/' });
                if (targetFolderPath.Length > 0)
                {
                    foreach (var folderName in targetFolderPath)
                    {
                        LogVerbose("Provisioning folder {0} into location {1}", folderName, relativeUri);
                        targetFolder = library.GetOrCreateFolder(targetFolder, folderName);
                        if(targetFolder == null)
                        {
                            LogWarning("Failed to provision folder {0}", folderName);
                            return false;
                        }
                    }
                }
                m_ensuredPath = true;
            }
            return m_ensuredPath;
        }


        /// <summary>
        /// Uploads a file in chunks if it exceeds the chunk size threshold
        /// </summary>
        /// <param name="m_clientContext">Context to the web</param>
        /// <param name="library">The target document library where it will be uploaded</param>
        /// <param name="webListFolderUrl">The target folder in the library</param>
        /// <param name="fileNameWithPath">Full path to the file</param>
        /// <param name="fileChunkSizeInMB">(OPTIONAL) defaults to 3 Megabytes</param>
        /// <returns></returns>
        public bool UploadFileWithBuffer(ClientContext m_clientContext, List library, Uri webListFolderUrl, string fileNameWithPath, int fileChunkSizeInMB = 3)
        {
            var folderPathDecoded = System.Web.HttpUtility.UrlDecode(webListFolderUrl.AbsolutePath);
            if (!EnsureFolderPath(m_clientContext, library, folderPathDecoded))
            {
                throw new Exception("Failed to ensure folder path directories.");
            }

            // Initialize File Info Properties
            var fileNameWithPathInfo = new FileInfo(fileNameWithPath);

            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file.
            string uniqueFileName = fileNameWithPathInfo.Name;


            var targetFolder = m_clientContext.Web.GetFolderByServerRelativeUrl(folderPathDecoded);
            m_clientContext.Load(targetFolder);
            m_clientContext.ExecuteQueryRetry();
            targetFolder.EnsureProperty(tfinc => tfinc.ServerRelativeUrl);
            var targetFolderUrl = TokenHelper.EnsureTrailingSlash(targetFolder.ServerRelativeUrl);


            // File object.
            Microsoft.SharePoint.Client.File uploadFile;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1048576;


            // Get the size of the file.
            var fileSize = fileNameWithPathInfo.Length;
            if (fileSize <= blockSize)
            {
                LogVerbose("File length {0}MB uploading with synchronous context {1}", fileSize.TryParseMB(), uniqueFileName);
                // Use regular approach.
                using (FileStream fs = new FileStream(fileNameWithPathInfo.FullName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation
                    {
                        ContentStream = fs,
                        Url = uniqueFileName,
                        Overwrite = true
                    };
                    uploadFile = targetFolder.Files.Add(fileInfo);
                    m_clientContext.Load(uploadFile);
                    m_clientContext.ExecuteQueryRetry();

                    // Return the file object for the uploaded file.
                    return true;
                }
            }
            else
            {
                var totalMBs = fileSize.TryParseMB();
                LogVerbose("File length {0}MB uploading with buffer context {1}", totalMBs, uniqueFileName);

                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileNameWithPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                LogVerbose("File length {0}MB uploading with synchronous context {1} => Started", totalMBs, uniqueFileName);

                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation
                                    {
                                        ContentStream = contentStream,
                                        Url = uniqueFileName,
                                        Overwrite = true
                                    };
                                    uploadFile = targetFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        m_clientContext.ExecuteQueryRetry();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                uploadFile = m_clientContext.Web.GetFileByServerRelativeUrl(targetFolderUrl + uniqueFileName);

                                if (last)
                                {
                                    LogVerbose("File length {0}MB uploading with synchronous context {1} FinishUploading => {2}", totalMBs, uniqueFileName, fileoffset.TryParseMB());

                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        m_clientContext.ExecuteQueryRetry();

                                        // Return the file object for the uploaded file.
                                        return true;
                                    }
                                }
                                else
                                {
                                    LogVerbose("File length {0}MB uploading with synchronous context {1} fileoffset => {2}MB", totalMBs, uniqueFileName, fileoffset.TryParseMB());

                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        m_clientContext.ExecuteQueryRetry();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }
            }

            return false;
        }

    }
}
