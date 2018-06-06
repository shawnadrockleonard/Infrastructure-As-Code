using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    /// <summary>
    /// The function cmdlet will migrate attachments from list items into a document library
    /// </summary>
    [Cmdlet(VerbsCommon.Move, "IaCAttachmentsToLibrary", SupportsShouldProcess = true)]
    public class MoveIaCAttachmentsToLibrary : IaCCmdlet
    {
        /// <summary>
        /// The source list containing attachments
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SourceListName { get; set; }

        /// <summary>
        /// The source url where documents will be copied
        /// </summary>
        [Parameter(Mandatory = false)]
        public string DestinationSiteUrl { get; set; }

        /// <summary>
        /// The source document library where files will be copied
        /// </summary>
        [Parameter(Mandatory = true)]
        public string DestinationLibraryName { get; set; }

        /// <summary>
        /// The action to take
        /// </summary>
        [Parameter(Mandatory = true)]
        [ValidateSet(new string[] { "create-folders", "copy-files" }, IgnoreCase = true)]
        public string SiteAction { get; set; }

        /// <summary>
        /// The starting path of the root folder
        /// </summary>
        [Parameter(Mandatory = false)]
        public string StartingFolder { get; set; }


        private System.IO.DirectoryInfo _resultLogDirectory { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            LogVerbose("Now migrating attachments from list to library {0}.", this.DestinationLibraryName);

            var result = System.IO.Path.GetTempPath();
            var resultdir = new System.IO.DirectoryInfo(result);
            _resultLogDirectory = resultdir.CreateSubdirectory("downloadedAttachments", resultdir.GetAccessControl());


            List _srclist = this.ClientContext.Web.Lists.GetByTitle(SourceListName);
            Folder _rootFolder = _srclist.RootFolder;
            this.ClientContext.Load(_srclist);
            this.ClientContext.Load(_rootFolder);
            this.ClientContext.ExecuteQuery();

            List _dlist = this.ClientContext.Web.Lists.GetByTitle(DestinationLibraryName);
            Folder _drootFolder = _dlist.RootFolder;
            this.ClientContext.Load(_dlist);
            this.ClientContext.Load(_drootFolder);
            this.ClientContext.ExecuteQuery();

            LogVerbose("Source List {0} with root folder {1}", _srclist.Title, _rootFolder.ServerRelativeUrl);
            LogVerbose("Destination List {0} with root folder {1}", _dlist.Title, _drootFolder.ServerRelativeUrl);

            //check if list is discussion list
            bool _isDiscussionList = false;
            ContentTypeCollection contentTypeColl = _srclist.ContentTypes;
            this.ClientContext.Load(contentTypeColl);
            this.ClientContext.ExecuteQuery();

            LogVerbose("Checking Content types:");
            foreach (ContentType contentType in contentTypeColl)
            {
                if (contentType.Name.IndexOf("Discussion") > -1)
                {
                    LogVerbose("Name: {0}  Id: {1}", contentType.Name, contentType.Id);
                    _isDiscussionList = true;
                    break;
                }
            }

            if (!(String.IsNullOrEmpty(StartingFolder)))
            {
                //Create StartingFolder
                _dlist.RootFolder.Folders.Add(StartingFolder);
                _dlist.Update();
                this.ClientContext.ExecuteQuery();
            }

            if (SiteAction.Equals("copy-files", StringComparison.CurrentCultureIgnoreCase))
            {
                //Process items for attachement
                GetItemsWithinFolder(_srclist, _dlist, _srclist.RootFolder);
            }


            if (!_isDiscussionList)
            {
                FolderCollection _folders = _rootFolder.Folders;
                this.ClientContext.Load(_folders);
                this.ClientContext.ExecuteQuery();

                foreach (Folder _folder in _folders)
                {
                    LogVerbose(">>>> {0}", _folder.Name);
                    if ((_folder.Name != "Attachments") && (_folder.Name != "Item"))
                    {
                        LogVerbose(">>>>" + _folder.Name);
                        ProcessFolder(_folder, _srclist, _dlist);
                    }
                }

            }
        }

        private void ProcessFolder(Folder _xfolder, List _list, List _dList)
        {
            LogVerbose("{0}--{1}", _xfolder.Name , _xfolder.ServerRelativeUrl);

            try
            {
                Folder _pFolder = _xfolder.ParentFolder;
                this.ClientContext.Load(_pFolder);
                this.ClientContext.ExecuteQuery();


                if (SiteAction.Equals("create-folders", StringComparison.CurrentCultureIgnoreCase))
                {
                    //Create folder in target library
                    CreateFolderInDocLib(_list, _dList, _xfolder, _pFolder);
                }

                if (SiteAction.Equals("copy-files", StringComparison.CurrentCultureIgnoreCase))
                {
                    //Process items for attachement
                    GetItemsWithinFolder(_list, _dList, _xfolder);
                }

                //Process subfolders
                FolderCollection _folders = _xfolder.Folders;
                this.ClientContext.Load(_folders);
                this.ClientContext.ExecuteQuery();

                foreach (Folder _folder in _folders)
                {
                    ProcessFolder(_folder, _list, _dList);
                }

            }
            catch (Exception e)
            {
                LogError(e,  ">>>>>> processFolder - {0}", e.Message);
            }
        }

        private void CreateFolderInDocLib(List _srcLst, List DocLib, Folder _folder, Folder _parentFolder)
        {

            string _newFolder = _folder.ServerRelativeUrl.Replace(_srcLst.RootFolder.ServerRelativeUrl, "");

            LogVerbose("----------------------------------------------------------------");
            LogVerbose("PREFOLDER: " + StartingFolder);
            LogVerbose("SrcList SrvRelUrl: " + _srcLst.RootFolder.ServerRelativeUrl);
            LogVerbose("Folder SrvRelUrl: " + _folder.ServerRelativeUrl);
            LogVerbose("New Folder1: " + _newFolder);

            _newFolder = _newFolder.Replace("//", "/").TrimEnd(new char[] { '/' });
            _newFolder = _newFolder.TrimStart(new char[] { '/' });
            _newFolder = _newFolder.Trim();


            Folder _folderX;


            LogVerbose("new Folder2: " + _newFolder);
            LogVerbose("DocLib Name: " + DocLib.RootFolder.Name);
            LogVerbose("DocLib Path: " + DocLib.RootFolder.ServerRelativeUrl);



            if ((_newFolder.IndexOf("/") > -1))
            {
                string[] _tmp = _newFolder.Split(new char[] { '/' });

                string _parentDstFolder = (_newFolder.Replace(_tmp[(_tmp.Length - 1)], "")).TrimEnd(new char[] { '/' });

                if (String.IsNullOrEmpty(StartingFolder))
                {
                    _parentDstFolder = DocLib.RootFolder.ServerRelativeUrl + "/" + _parentDstFolder + "/";
                }
                else
                {
                    _parentDstFolder = DocLib.RootFolder.ServerRelativeUrl + "/" + StartingFolder + "/" + _parentDstFolder + "/";
                }


                _parentDstFolder = _parentDstFolder.Replace("//", "/");

                LogVerbose("--->> {0} ++>> {1}", _parentDstFolder, _tmp[(_tmp.Length - 1)]);


                _parentDstFolder = _parentDstFolder.TrimEnd(new char[] { '/' });

                Folder _parentDstFolderObj = this.ClientContext.Web.GetFolderByServerRelativeUrl(_parentDstFolder);
                this.ClientContext.Load(_parentDstFolderObj);
                this.ClientContext.ExecuteQuery();


                Folder _folderXdst = _parentDstFolderObj.Folders.Add(_tmp[(_tmp.Length - 1)]);
                this.ClientContext.Load(_folderXdst);
                DocLib.Update();
                this.ClientContext.ExecuteQuery();
            }
            else
            {

                var _newFolder2 = _folder.ServerRelativeUrl.Replace(_srcLst.RootFolder.ServerRelativeUrl, "");

                var _tmp = _newFolder2.Split(new char[] { '/' });

                var _parentDstFolder = (_newFolder2.Replace(_tmp[(_tmp.Length - 1)], "")).TrimEnd(new char[] { '/' });

                if (String.IsNullOrEmpty(StartingFolder))
                {
                    _parentDstFolder = DocLib.RootFolder.ServerRelativeUrl + "/" + _parentDstFolder + "/";
                }
                else
                {
                    _parentDstFolder = DocLib.RootFolder.ServerRelativeUrl + "/" + StartingFolder + "/" + _parentDstFolder + "/";
                }


                _parentDstFolder = _parentDstFolder.Replace("//", "/");

                LogVerbose("--000->> {0} +00+>> {1}", _parentDstFolder, _tmp[(_tmp.Length - 1)]);

                if (String.IsNullOrEmpty(StartingFolder))
                {
                    LogVerbose("---------------");
                    _folderX = DocLib.RootFolder.Folders.Add(_newFolder);
                    this.ClientContext.Load(_folderX);
                    // _parentFolder.Update();
                    DocLib.Update();
                    this.ClientContext.ExecuteQuery();

                }
                else
                {
                    Folder _parentDstFolderObj = this.ClientContext.Web.GetFolderByServerRelativeUrl(_parentDstFolder);
                    this.ClientContext.Load(_parentDstFolderObj);
                    this.ClientContext.ExecuteQuery();

                    LogVerbose("Parent url: " + _parentDstFolderObj.ServerRelativeUrl);

                    Folder _folderXdst = _parentDstFolderObj.Folders.Add(_tmp[(_tmp.Length - 1)]);
                    this.ClientContext.Load(_folderXdst);
                    DocLib.Update();
                    this.ClientContext.ExecuteQuery();
                }


            }


            this.ClientContext.ExecuteQuery();
        }

        private void GetItemsWithinFolder(List _srcLst, List _dList, Folder _folder)
        {
            CamlQuery camlQuery = new CamlQuery();
            camlQuery = new CamlQuery();
            camlQuery.FolderServerRelativeUrl = _folder.ServerRelativeUrl;
            camlQuery.ViewXml = @"<View><RowLimit Paged='TRUE'>30</RowLimit></View>";
            ListItemCollection listItems = _srcLst.GetItems(camlQuery);
            this.ClientContext.Load(listItems, itl => itl.Include(itln => itln.Id, itln => itln.DisplayName));
            this.ClientContext.ExecuteQuery();

            foreach (ListItem _item in listItems)
            {
                LogVerbose("--->> {0} -- {1}", +_item.Id, _item.DisplayName);
                LogVerbose("_dList.RootFolder.ServerRelativeUrl: " + _dList.RootFolder.ServerRelativeUrl);
                LogVerbose("_srcLst.RootFolder.ServerRelativeUrl: " + _srcLst.RootFolder.ServerRelativeUrl);
                LogVerbose("+--+ >> " + _folder.ServerRelativeUrl);
                CopyAttachment(_srcLst, _dList, _item.Id, _folder.ServerRelativeUrl);
            }
        }

        private void CopyAttachment(List _srcLst, List _dList, int _itemID, string _folder)
        {
            string _tmpfolder = "";

            if (String.IsNullOrEmpty(StartingFolder))
            {
                _tmpfolder = _folder.Replace((_srcLst.RootFolder.ServerRelativeUrl), _dList.RootFolder.ServerRelativeUrl);
            }
            else
            {
                string _tmpSeed = _dList.RootFolder.ServerRelativeUrl + "/" + StartingFolder;
                _tmpfolder = _folder.Replace((_srcLst.RootFolder.ServerRelativeUrl), _tmpSeed);
            }

            LogVerbose("Copying: Items in " + _itemID + " To " + _tmpfolder);
            LogVerbose("Source file: " + _srcLst.RootFolder.ServerRelativeUrl + "/Attachments/" + _itemID);

            try
            {
                Folder _aFolder = this.ClientContext.Web.GetFolderByServerRelativeUrl(_srcLst.RootFolder.ServerRelativeUrl + "/Attachments/" + _itemID);
                try
                {
                    this.ClientContext.Load(_aFolder);
                    this.ClientContext.ExecuteQuery();
                }
                catch (ServerException ex)
                {
                    LogError(ex, "No attachments for ListItem {0}", ex.Message);
                    LogWarning("No Attachment for ID {0}", _itemID);
                    return;
                }


                FileCollection _files = _aFolder.Files;
                this.ClientContext.Load(_files);
                this.ClientContext.ExecuteQuery();

                foreach (var _file in _files)
                {
                    try
                    {
                        string _tmpFileName = HelperExtensions.GetCleanFileName(_file.Name);
                        string url = _tmpfolder + "/" + _tmpFileName;

                        LogVerbose("Filename: " + _file.Name);
                        if (this.ShouldProcess(string.Format("Will upload local file {0} to url {1}", _file.Name, url)))
                        {
                            LogVerbose("_tmpfolder: {0} _tmpFileName: {1} with new file url: {2}" + _tmpfolder, _tmpFileName, url);
                            var fileRelativeUrl = _file.ServerRelativeUrl.ToString();
                            var fileNameFromInfo = System.IO.Path.GetFileName(fileRelativeUrl);
                            var doclibfilename = string.Format("{0}\\{1}", _resultLogDirectory, fileNameFromInfo);

                            // read file from attachment pool
                            FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(this.ClientContext, fileRelativeUrl);
                            this.ClientContext.ExecuteQueryRetry();
                            
                            // write binary to file
                            using (var fileStream = new System.IO.FileStream(doclibfilename, System.IO.FileMode.Create))
                            {
                                fileInfo.Stream.CopyTo(fileStream);
                            }

                            // read file from disk into memorystream and upload
                            using (var ms = new System.IO.MemoryStream(System.IO.File.ReadAllBytes(doclibfilename)))
                            {
                                Microsoft.SharePoint.Client.File.SaveBinaryDirect(this.ClientContext, url, ms, false);
                                this.ClientContext.ExecuteQuery();
                            }

                            Thread.Sleep(500);
                        }

                    }
                    catch (Exception e)
                    {
                        LogError(e,  ">>>>>> {0}", e.Message);
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e,  ">>>>>> copyAttachment - {0}", e.Message);
            }
        }

    }
}
