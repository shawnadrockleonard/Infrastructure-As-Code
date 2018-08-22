using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    /// <summary>
    /// The function cmdlet will watch a directory and upload the files
    /// </summary>
    [Cmdlet(VerbsCommon.Watch, "IaCDirectoryAndUpload", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low)]
    public class WatchIaCDirectoryAndUpload : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SiteContent { get; set; }

        [Parameter(Mandatory = true)]
        public ListPipeBind TargetList { get; set; }

        /// <summary>
        /// The single SiteAsset file to upload based on relative path
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "FileAction")]
        public string SiteActionFile { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "FileWatcher")]
        public SwitchParameter Watch { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "FileWatcher")]
        public string[] FileNameFilters { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "FileWatcher")]
        public int TestSeconds = 5;

        [Parameter(Mandatory = false, ParameterSetName = "FileWatcher")]
        public int WaitSeconds = 5;

        /// <summary>
        /// Represents the initial seeded datetime to compare files
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "FileWatcher")]
        public Nullable<DateTime> CompareDatetime { get; set; }


        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!System.IO.Directory.Exists(this.SiteContent))
            {
                throw new Exception(string.Format("The directory does not exists {0}", this.SiteContent));
            }

            SiteContentDirectory = new System.IO.DirectoryInfo(this.SiteContent);
        }

        private System.IO.DirectoryInfo SiteContentDirectory { get; set; }

        private DateTime _LastWriteTime { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // obtain CSOM object for host web
            Web hostWeb = this.ClientContext.Web;

            if (!CompareDatetime.HasValue)
            {
                _LastWriteTime = DateTime.Now;
            }
            else
            {
                _LastWriteTime = CompareDatetime.Value;
            }


            // check to see if library exists
            var siteLibrary = TargetList.GetList(this.ClientContext.Web);
            if (siteLibrary == null)
            {
                LogWarning("Failed to find site list/library {0}", TargetList.ToString());
                return;
            }


            if (Watch)
            {
                try
                {
                    siteLibrary.EnsureProperties(sl => sl.RootFolder, sl => sl.RootFolder.ServerRelativeUrl);

                    var watcherFilters = (FileNameFilters == null || !FileNameFilters.Any()) ? new string[] { "*.*" } : FileNameFilters;
                    var watcherFiles = watcherFilters.SelectMany(sm => System.IO.Directory.GetFiles(this.SiteContentDirectory.FullName, sm, System.IO.SearchOption.AllDirectories));

                    while (true)
                    {
                        System.Threading.Thread.Sleep(new TimeSpan(0, 0, TestSeconds));
                        if (this.Stopping)
                        {
                            LogWarning("Stopping the process and discontinuing the watcher");
                            break;
                        }
                        var changedItems = watcherFiles
                            .Select(sf => new System.IO.FileInfo(sf))
                            .Where(wf =>
                            {
                                return ((wf.LastWriteTime.Subtract(_LastWriteTime)).TotalSeconds >= WaitSeconds);
                            })
                            .OrderBy(ob => ob.DirectoryName)
                            .ThenBy(tb => tb.LastWriteTime)
                            .ToList();

                        if (changedItems != null && changedItems.Any())
                        {
                            _LastWriteTime = DateTime.Now;
                            var parentName = string.Empty;
                            var tmpParentname = string.Empty;
                            var sitecontentfullname = SiteContentDirectory.FullName;
                            Folder directoryPath = null;

                            foreach (var change in changedItems)
                            {
                                parentName = change.Directory.Name;
                                if (parentName != tmpParentname)
                                {
                                    var filedirectorypath = change.Directory.FullName.Replace(sitecontentfullname, "").Replace("\\", "/");
                                    directoryPath = siteLibrary.RootFolder.ListEnsureFolder(filedirectorypath);
                                    tmpParentname = parentName;
                                }

                                OnChanged(siteLibrary, directoryPath, change);
                            }
                        }
                    }
                }
                catch (Exception tex)
                {
                    LogWarning("Terminating watch command {0}", tex.Message);
                }
            }
            else
            {
                var searchPattern = "*";
                if (!string.IsNullOrEmpty(this.SiteActionFile))
                {
                    searchPattern = this.SiteActionFile;
                    if (this.SiteActionFile.IndexOf(@"\") > -1)
                    {
                        searchPattern = this.SiteActionFile.Substring(0, this.SiteActionFile.IndexOf(@"\"));
                    }
                }

                var appDirectories = System.IO.Directory.GetDirectories(this.SiteContent, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
                foreach (var appDirectory in appDirectories)
                {
                    var appDirectoryInfo = new System.IO.DirectoryInfo(appDirectory);
                    UploadSiteAssetFilesToWeb(siteLibrary, siteLibrary.RootFolder, appDirectoryInfo);
                }
            }
        }


        protected override void StopProcessing()
        {
            base.StopProcessing();
            System.Diagnostics.Trace.TraceWarning("Stopping the pipeline");
        }

        private void OnChanged(List siteAssetsLibrary, Folder parentFolder, System.IO.FileInfo source)
        {
            //Copies file to another directory.
            ShouldProcessReason _process;
            var filePath = source.FullName;
            var fileName = System.IO.Path.GetFileName(filePath);
            if (this.ShouldProcess(
                string.Format("Uploading {0} timestamp {1}", fileName, source.LastWriteTime), 
                string.Format("Uploading {0} timestamp {1}", fileName, source.LastWriteTime), 
                string.Format("Caption {0}", fileName), out _process))
            {
                LogVerbose("---------------- Now uploading file {0}", filePath);
                // upload each file to library in host web
                byte[] fileContent = System.IO.File.ReadAllBytes(filePath);
                FileCreationInformation fileInfo = new FileCreationInformation();
                fileInfo.Content = fileContent;
                fileInfo.Overwrite = true;
                fileInfo.Url = fileName;
                File newFile = parentFolder.Files.Add(fileInfo);

                // commit changes to library
                this.ClientContext.Load(newFile, nf => nf.ServerRelativeUrl, nf => nf.Length);
                this.ClientContext.ExecuteQueryRetry();
            }
        }

        private void UploadSiteAssetFilesToWeb(List siteAssetsLibrary, Folder parentfolder, System.IO.DirectoryInfo destinationFolder)
        {
            var folderPath = destinationFolder.FullName;
            var searchPattern = "*";
            if (!string.IsNullOrEmpty(this.SiteActionFile))
            {
                if (this.SiteActionFile.IndexOf(@"\") > -1)
                {
                    searchPattern = string.Format("*{0}*", this.SiteActionFile.Substring(this.SiteActionFile.IndexOf(@"\") + 1));
                }
            }

            var folder = parentfolder.EnsureFolder(destinationFolder.Name);

            var siteAssetsFiles = System.IO.Directory.GetFiles(folderPath, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
            LogVerbose("Now searching folder {0} and uploading files {1}", folderPath, siteAssetsFiles.Count());

            // enmuerate through each file in folder
            foreach (string filePath in siteAssetsFiles)
            {
                OnChanged(siteAssetsLibrary, folder, new System.IO.FileInfo(filePath));
            }

            var siteAssetsFolders = System.IO.Directory.GetDirectories(folderPath, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
            if (siteAssetsFolders.Count() > 0)
            {
                siteAssetsFolders.ToList().ForEach(subFolderPath =>
                {
                    var appDirectoryInfo = new System.IO.DirectoryInfo(subFolderPath);
                    UploadSiteAssetFilesToWeb(siteAssetsLibrary, folder, appDirectoryInfo);
                });
            }

        }

    }
}
