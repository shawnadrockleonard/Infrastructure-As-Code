using CsvHelper;
using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    /// <summary>
    /// Uploads a document or file to a library specified 
    ///     Metadatalist is a CSV file containing metadata associated with the Filename
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCASyncDirectory")]
    public class SetIaCSyncDirectory : IaCCmdlet
    {
        /// <summary>
        /// Represents the display title for the document library
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string LibraryName { get; set; }

        /// <summary>
        /// Directory where the documents to be uploaded exist
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string DirectoryPath { get; set; }

        /// <summary>
        /// CSV file which contains metadata
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string MetaDataCsvFile { get; set; }

        /// <summary>
        /// Should the process create the folders and upload or just create folders
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter UploadFiles { get; set; }

        /// <summary>
        /// Does the library require check-in process
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter CheckIn { get; set; }

        /// <summary>
        /// Represents a collection of metadata elements
        /// </summary>
        internal List<FileTagModel> MetadataList { get; set; }

        /// <summary>
        /// Check defaults and processing requirements
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (!System.IO.Directory.Exists(this.DirectoryPath))
            {
                throw new InvalidOperationException(String.Format("{0} directory does not exist or this program can not find it.", this.DirectoryPath));
            }

            this.MetadataList = new List<FileTagModel>();
        }

        /// <summary>
        /// Process the user request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            try
            {
                if (System.IO.File.Exists(this.MetaDataCsvFile))
                {
                    using (var fileInfo = System.IO.File.OpenText(this.MetaDataCsvFile))
                    {
                        var csv = new CsvReader(fileInfo);
                        csv.Configuration.HasHeaderRecord = true;
                        csv.Configuration.Delimiter = ",";
                        csv.Configuration.RegisterClassMap<FileTagModelMap>();

                        this.MetadataList.AddRange(csv.GetRecords<FileTagModel>());
                    }
                }
            }
            catch (System.IO.FileNotFoundException fex)
            {
                LogError(fex, "Failed to reach CSV file");
            }

            try
            {
                var onlineLibrary = this.ClientContext.Web.Lists.GetByTitle(this.LibraryName);
                this.ClientContext.Load(onlineLibrary, ol => ol.RootFolder, ol => ol.Title, ol => ol.EnableVersioning);
                this.ClientContext.ExecuteQuery();

                if (onlineLibrary.EnableVersioning)
                {
                    onlineLibrary.UpdateListVersioning(false, false, true);
                }

                // Parse Top Directory, enumerate files and upload them to root folder
                if (this.UploadFiles)
                {
                    UploadFileToSharePointFolder(onlineLibrary.RootFolder, this.DirectoryPath);
                }

                // Loop through all folders (recursive) that exist within the folder supplied by the operator
                var firstLevelFolders = System.IO.Directory.GetDirectories(this.DirectoryPath, "*", System.IO.SearchOption.TopDirectoryOnly);
                foreach (var folder in firstLevelFolders)
                {
                    var createdFolder = PopulateSharePointFolderWithFiles(onlineLibrary.RootFolder, this.DirectoryPath, folder);
                    if (!createdFolder)
                    {
                        LogDebugging("Folder {0} was not created successfully.", folder);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in SetDirectoryUpload {0} directory", DirectoryPath);
            }

        }

        /// <summary>
        /// Recursively calls the folder creation method for the folder
        /// </summary>
        /// <param name="parentFolder">Set the Folder equal to the absolute path of the SharePoint destination</param>
        /// <param name="directoryPath">The directory path which includes the fullFolderUrl</param>
        /// <param name="fullFolderUrl"></param>
        /// <returns></returns>
        private bool PopulateSharePointFolderWithFiles(Folder parentFolder, string directoryPath, string fullFolderUrl)
        {
            var subStatus = false;
            try
            {

                //Set the FolderRelativePath by removing the path of the folder supplied by the operator from the fullname of the folder
                var folderInfo = new System.IO.DirectoryInfo(fullFolderUrl);
                var folderRelativePath = folderInfo.FullName.Substring(directoryPath.Length); // should filter out the parent folder
                if (folderRelativePath.StartsWith(@"\"))
                {
                    folderRelativePath = folderRelativePath.Substring(1);
                }

                string trimmedFolder = folderRelativePath.Trim().Replace("_", " "); // clean the folder name for SharePoint purposes

                // setup processing of folder in the parent folder
                var currentFolder = parentFolder;
                this.ClientContext.Load(parentFolder, pf => pf.Name, pf => pf.Folders, pf => pf.Files);
                this.ClientContext.ExecuteQuery();

                if (!parentFolder.FolderExists(trimmedFolder))
                {
                    currentFolder = parentFolder.EnsureFolder(trimmedFolder);
                    //this.ClientContext.Load(curFolder);
                    this.ClientContext.ExecuteQuery();
                    LogVerbose(".......... successfully created folder {0}....", folderRelativePath);
                }
                else
                {
                    currentFolder = parentFolder.Folders.FirstOrDefault(f => f.Name == trimmedFolder);
                    LogVerbose(".......... reading folder {0}....", folderRelativePath);
                }

                // Powershell parameter switch
                if (this.UploadFiles)
                {
                    UploadFileToSharePointFolder(currentFolder, folderInfo.FullName);
                }


                // retrieve any subdirectories for the child folder
                var firstLevelFolders = folderInfo.EnumerateDirectories().ToList();
                if (firstLevelFolders.Count() > 0)
                {
                    LogVerbose("Creating folders for {0}.....discovering folders {1}", folderRelativePath, firstLevelFolders.Count());

                    foreach (var subFolderUrl in firstLevelFolders)
                    {
                        subStatus = PopulateSharePointFolderWithFiles(currentFolder, fullFolderUrl, subFolderUrl.FullName);
                    }

                    LogVerbose("Leaving folder {0}.....", folderRelativePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to provision SPFolders in SPFolder:{0}", fullFolderUrl);
            }
            return false;
        }

        /// <summary>
        /// Upload the specific file to the destination folder
        /// </summary>
        /// <param name="checkIn"></param>
        /// <param name="destinationFolder"></param>
        /// <param name="fileNameWithPath"></param>
        /// <returns></returns>
        private bool UploadFileToSharePointFolder(Microsoft.SharePoint.Client.Folder destinationFolder, string directoryPath)
        {
            // Retrieve the files in the directory
            var filesInDirectPath = System.IO.Directory.GetFiles(directoryPath, "*", System.IO.SearchOption.TopDirectoryOnly);
            if (filesInDirectPath.Count() > 0)
            {
                //  Upload the file
                LogVerbose("Uploading directory {0} files", directoryPath);

                //  For each file in the source folder being evaluated, call the UploadFile function to upload the file to the appropriate location
                foreach (var fileNameWithPath in filesInDirectPath)
                {
                    try
                    {
                        var fileExists = false;
                        var fileTags = string.Empty;
                        var fileInfo = new System.IO.FileInfo(fileNameWithPath);

                        //  Notify the operator that the file is being uploaed to a specific location
                        LogVerbose("Uploading file {0} to {1}", fileInfo.Name, destinationFolder.Name);

                        try
                        {
                            var fileInFolder = destinationFolder.GetFile(fileInfo.Name);
                            if (fileInFolder != null)
                            {
                                fileExists = true;
                                LogVerbose("File {0} exists in the destination folder.  Skip uploading file.....", fileInfo.Name);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogError(ex, "Failed check file {0} existance test.", fileInfo.Name);
                        }

                        if (!fileExists)
                        {
                            try
                            {
                                if (this.MetadataList.Any(md => md.FullPath == fileInfo.FullName))
                                {
                                    var distinctTags = this.MetadataList.Where(md => md.FullPath == fileInfo.FullName).Select(s => s.Tag).Distinct();
                                    fileTags = string.Join(@";", distinctTags);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogDebugging("Failed to pull metadata from CSV file {0}", ex.Message);
                            }

                            using (var stream = new System.IO.FileStream(fileNameWithPath, System.IO.FileMode.Open))
                            {

                                var creationInfo = new Microsoft.SharePoint.Client.FileCreationInformation();
                                creationInfo.Overwrite = true;
                                creationInfo.ContentStream = stream;
                                creationInfo.Url = fileInfo.Name;

                                var uploadStatus = destinationFolder.Files.Add(creationInfo);
                                if (!uploadStatus.IsPropertyAvailable("ListItemAllFields"))
                                {
                                    this.ClientContext.Load(uploadStatus, w => w.ListItemAllFields);
                                    this.ClientContext.ExecuteQuery();
                                }
                                if (!string.IsNullOrEmpty(fileTags))
                                {
                                    uploadStatus.ListItemAllFields["_Source"] = fileTags;
                                }
                                uploadStatus.ListItemAllFields[ConstantsListFields.Field_Modified] = fileInfo.LastWriteTime;
                                uploadStatus.ListItemAllFields.SystemUpdate();

                                if (this.CheckIn)
                                {
                                    this.ClientContext.Load(uploadStatus);
                                    this.ClientContext.ExecuteQuery();
                                    if (uploadStatus.CheckOutType != Microsoft.SharePoint.Client.CheckOutType.None)
                                    {
                                        uploadStatus.CheckIn("Checked in by Administrator", Microsoft.SharePoint.Client.CheckinType.MajorCheckIn);
                                    }
                                }
                                this.ClientContext.Load(uploadStatus);
                                this.ClientContext.ExecuteQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError(ex, "Failed in UploadFileToSharePointFolder destination {0} and file {1}", destinationFolder, fileNameWithPath);
                    }
                }

                return true; // Has Files
            }
            return false; // No Files in Directory
        }
    }
}
