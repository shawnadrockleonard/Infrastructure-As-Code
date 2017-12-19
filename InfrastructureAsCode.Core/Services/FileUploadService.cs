using InfrastructureAsCode.Core.Reports;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InfrastructureAsCode.Core.Utilities;

namespace InfrastructureAsCode.Core.Services
{
    public class FileUploadService
    {
        /// <summary>
        /// Initializes the file upload control
        /// </summary>
        /// <param name="SharePointClientContext">The client context required to process CSOM requests</param>
        /// <param name="DiagnosticLogger">The diagnostic logger for storing in your system</param>
        public FileUploadService(ClientContext SharePointClientContext, ITraceLogger DiagnosticLogger)
        {
            m_clientContext = SharePointClientContext;
            logger = DiagnosticLogger;
        }

        #region Private variables 

        /// <summary>
        /// Context to the web
        /// </summary>
        private ClientContext m_clientContext { get; set; }

        /// <summary>
        /// The diagnostic logger for writing diagnostics to a log agent or the screen
        /// </summary>
        private ITraceLogger logger { get; set; }

        #endregion

        /// <summary>
        /// Will accept the web server relative path and ensure the folder structure is added to the <paramref name="library"/>
        /// </summary>
        /// <param name="library">The document library</param>
        /// <param name="webListFolderUrl">A web server relative URL including the folder structure where the file will be uploaded</param>
        /// <returns></returns>
        private bool EnsureFolderPath(List library, string webListFolderUrl)
        {
            // make sure the RootFolder object is loaded into memory
            if (!library.IsPropertyAvailable(lctx => lctx.RootFolder)
                || library.IsPropertyAvailable(lctx => lctx.RootFolder.ServerRelativeUrl))
            {
                library.EnsureProperties(lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl);
            }

            var targetFolder = m_clientContext.Web.GetFolderByServerRelativeUrl(webListFolderUrl);
            try
            {
                m_clientContext.Load(targetFolder);
                m_clientContext.ExecuteQueryRetry();
                return true;
            }
            catch (Exception ex)
            {
                logger.LogWarning("Folder check exception {0}", ex.Message);

                // Folder Not Found
                // Enumerate the Folder structure creating the desired destination 
                targetFolder = library.RootFolder;
                var relativeUri = TokenHelper.EnsureTrailingSlash(targetFolder.ServerRelativeUrl);
                var targetFolderPath = webListFolderUrl.Replace(relativeUri, string.Empty).Split(new char[] { '/' });
                if (targetFolderPath.Length > 0)
                {
                    foreach (var folderName in targetFolderPath)
                    {
                        logger.LogInformation("Provisioning folder {0} into location {1}", folderName, relativeUri);
                        targetFolder = library.GetOrCreateFolder(targetFolder, folderName);
                        if (targetFolder == null)
                        {
                            logger.LogWarning("Failed to provision folder {0}", folderName);
                            return false;
                        }
                    }
                }
            }

            return true;
        }


        /// <summary>
        /// Uploads a file in chunks if it exceeds the chunk size threshold
        /// </summary>
        /// <param name="m_clientContext"></param>
        /// <param name="library">The target document library where it will be uploaded</param>
        /// <param name="webListFolderUrl">The target folder in the library</param>
        /// <param name="fileNameWithPath">Full path to the file</param>
        /// <param name="fileChunkSizeInMB">(OPTIONAL) defaults to 3 Megabytes</param>
        /// <returns></returns>
        public bool UploadFileWithBuffer(List library, Uri webListFolderUrl, string fileNameWithPath, int fileChunkSizeInMB = 3)
        {

            // Initialize File Info Properties
            var fileNameWithPathInfo = new System.IO.FileInfo(fileNameWithPath);

            // Get the size of the file.
            var fileSize = fileNameWithPathInfo.Length;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1048576;

            // Validate the chunk size and ensure its lower than 10 Mbs
            if (fileSize > blockSize && fileChunkSizeInMB > 10)
            {
                throw new ArgumentException(string.Format("Your file chunk size is set too high {0}", fileChunkSizeInMB));
            }


            var folderPathDecoded = System.Web.HttpUtility.UrlDecode(webListFolderUrl.AbsolutePath);
            if (!EnsureFolderPath(library, folderPathDecoded))
            {
                throw new ArgumentException("Failed to ensure folder path directories.");
            }


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



            if (fileSize <= blockSize)
            {
                logger.LogInformation("File length {0}MB uploading with synchronous context {1}", fileSize.TryParseMB(), uniqueFileName);
                // Use regular approach.
                using (var fs = new System.IO.FileStream(fileNameWithPathInfo.FullName, System.IO.FileMode.Open))
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
                logger.LogInformation("File length {0}MB uploading with buffer context {1}", totalMBs, uniqueFileName);

                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                System.IO.FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileNameWithPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                    using (var br = new System.IO.BinaryReader(fs))
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
                                logger.LogInformation("File length {0}MB uploading with synchronous context {1} => Started", totalMBs, uniqueFileName);

                                using (var contentStream = new System.IO.MemoryStream())
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
                                    using (var s = new System.IO.MemoryStream(buffer))
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
                                    logger.LogInformation("File length {0}MB uploading with synchronous context {1} FinishUploading => {2}", totalMBs, uniqueFileName, fileoffset.TryParseMB());

                                    // Is this the last slice of data?
                                    using (var s = new System.IO.MemoryStream(lastBuffer))
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
                                    logger.LogInformation("File length {0}MB uploading with synchronous context {1} fileoffset => {2}MB", totalMBs, uniqueFileName, fileoffset.TryParseMB());

                                    using (var s = new System.IO.MemoryStream(buffer))
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
