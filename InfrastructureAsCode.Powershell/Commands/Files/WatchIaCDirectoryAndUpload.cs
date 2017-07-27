using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
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
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!System.IO.Directory.Exists(this.SiteContent))
            {
                throw new Exception(string.Format("The directory does not exists {0}", this.SiteContent));
            }
        }


        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // obtain CSOM object for host web
            Web hostWeb = this.ClientContext.Web;


            // check to see if library exists
            var listTitle = TargetList.Title;
            ListCollection allLists = hostWeb.Lists;
            IEnumerable<List> foundLists = this.ClientContext.LoadQuery(allLists.Where(list => list.Title == listTitle));
            this.ClientContext.ExecuteQueryRetry();
            var siteLibrary = foundLists.FirstOrDefault();
            if (siteLibrary == null)
            {
                LogWarning("Failed to find site library {0}", TargetList.Title);
                return;
            }


            if (Watch)
            {
                try
                {
                    var watcherFilters = (FileNameFilters == null || !FileNameFilters.Any()) ? new string[] { "*.*" } : FileNameFilters;
                    var watcherFiles = watcherFilters.SelectMany(sm => System.IO.Directory.GetFiles(this.SiteContent, sm, System.IO.SearchOption.AllDirectories));


                    DateTime lastTime = DateTime.Now;
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
                                return ((wf.LastWriteTime - lastTime).TotalSeconds >= WaitSeconds);
                            })
                            .OrderBy(ob => ob.DirectoryName)
                            .ThenBy(tb => tb.LastWriteTime)
                            .ToList();

                        if (changedItems != null && changedItems.Any())
                        {
                            lastTime = DateTime.Now;
                            var parentName = string.Empty;
                            var tmpParentname = string.Empty;
                            var fileFolder = siteLibrary.RootFolder;

                            foreach (var change in changedItems)
                            {
                                parentName = change.Directory.Name;
                                if (parentName != tmpParentname)
                                {
                                    fileFolder = siteLibrary.RootFolder.EnsureFolder(parentName);
                                    tmpParentname = parentName;
                                }

                                OnChanged(siteLibrary, fileFolder, change);
                            }
                        }
                    }
                }
                catch (System.Management.Automation.TerminateException tex)
                {
                    LogWarning("Terminating watch command {0}", tex.Message);
                }
            }
            else
            {
                var appFileFolder = string.Format("{0}\\SiteAssets", this.SiteContent);
                if (!System.IO.Directory.Exists(appFileFolder))
                {
                    LogWarning("Site Assets Folder {0} does not exist.", appFileFolder);
                    return;
                }

                var searchPattern = "*";
                if (!string.IsNullOrEmpty(this.SiteActionFile))
                {
                    searchPattern = this.SiteActionFile;
                    if (this.SiteActionFile.IndexOf(@"\") > -1)
                    {
                        searchPattern = this.SiteActionFile.Substring(0, this.SiteActionFile.IndexOf(@"\"));
                    }
                }

                var appDirectories = System.IO.Directory.GetDirectories(appFileFolder, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
                foreach (var appDirectory in appDirectories)
                {
                    var appDirectoryInfo = new System.IO.DirectoryInfo(appDirectory);
                    UploadSiteAssetFilesToWeb(siteLibrary, siteLibrary.RootFolder, appDirectoryInfo);
                }
            }
        }



        private void OnChanged(List siteAssetsLibrary, Folder parentFolder, System.IO.FileInfo source)
        {
            //Copies file to another directory.
            var filePath = source.FullName;
            var fileName = System.IO.Path.GetFileName(filePath);
            if (this.ShouldProcess(string.Format("Should upload {0} timestamp {1}", fileName, source.LastWriteTime)))
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
