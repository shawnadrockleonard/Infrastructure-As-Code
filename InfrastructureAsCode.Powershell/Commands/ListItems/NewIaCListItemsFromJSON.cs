using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Web.Script.Serialization;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    /// <summary>
    /// Demonstrates how to create specific items based on the JSON specified. 
    ///     
    /// </summary>
    [Cmdlet(VerbsCommon.New, "IaCListItemsFromJSON", SupportsShouldProcess = true)]
    [CmdletHelp("Query the JSON file and insert new rows into the target list.", Category = "ListItems")]
    public class NewIaCListItemsFromJSON : IaCCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "The fully qualified path to the JSON file.")]
        public string JSONFile { get; set; }

        [Parameter(Mandatory = false,  Position = 2, HelpMessage = "The fully qualified path to export the C# class file.")]
        public string CSharpFile { get; set; }

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (!(new System.IO.FileInfo(JSONFile)).Exists)
            {
                throw new System.IO.FileNotFoundException($"Could not find the JSON {JSONFile} file specified");
            }

            if (ClientContext == null)
            {
                throw new Exception($"The specified context is null; please reconnect");
            }
        }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var siteUrl = ClientContext.Url;
            var dtStart = DateTime.Now;
            var strJSON = string.Empty;

            // Read file
            using (StreamReader sr = new StreamReader(JSONFile))
            {
                strJSON = sr.ReadToEnd();
            }

            // Process the JSON
            var jsSerializer = new JavaScriptSerializer();
            SharePointContents spContents = (SharePointContents)jsSerializer.Deserialize(strJSON, typeof(SharePointContents));


            foreach (SharePointList spList in spContents.SharePointLists)
            {
                if (!spList.Ignore)
                {
                    // delete list, using new context
                    using (var cContext = this.ClientContext.Clone(siteUrl))
                    {
                        DeleteList(cContext, spList.Name);
                    }

                    // add list, using new context
                    using (var cContext = this.ClientContext.Clone(siteUrl))
                    {
                        CreateList(cContext, spList.Name, spList.Type);
                    }

                    // modify list and data, using new context
                    using (var cContext = this.ClientContext.Clone(siteUrl))
                    {
                        if (spList.TitleName != null && spList.TitleName != string.Empty)
                        {
                            if (spList.TitleName != "DELETE")
                            {
                                RenameTitle(cContext, spList);
                            }
                            else
                            {
                                DeleteTitle(cContext, spList);
                            }
                        }
                        else
                        {
                            spList.TitleName = "Title";
                        }

                        foreach (SharePointColumn spColumn in spList.SharePointColumns)
                        {
                            AddColumn(cContext, spColumn.Name, spColumn.Type, spList.Name, spColumn.Parent, spColumn.Data);
                        }

                        AddReferenceData(cContext, spList);
                    }
                }
            }

            //Optional, creates c# class
            if (!string.IsNullOrEmpty(CSharpFile))
            {
                GenerateCSharp(spContents);
            }

            LogVerbose($"DONE! {(DateTime.Now - dtStart).TotalSeconds.ToString()} seconds");
        }

        internal void DeleteTitle(ClientContext cContext, SharePointList spList)
        {
            try
            {
                LogVerbose("Start DeleteTitle: " + spList.Name);
                Web wWeb = cContext.Web;
                List lList = wWeb.Lists.GetByTitle(spList.Name);
                FieldCollection collField = lList.Fields;
                Field oneField = collField.GetByInternalNameOrTitle("Title");

                oneField.Hidden = true;
                oneField.Required = false;
                oneField.SetShowInDisplayForm(false);
                oneField.SetShowInEditForm(false);
                oneField.SetShowInNewForm(false);

                View view = lList.Views.GetByTitle("All Items");
                ViewFieldCollection viewFields = view.ViewFields;
                viewFields.Remove("LinkTitle");
                view.Update();

                oneField.Update();
                cContext.ExecuteQuery();
                LogVerbose("Complete DeleteTitle: " + spList.Name);
            }
            catch (Exception ex)
            {
                LogError(ex, ex.Message);
            }
        }

        /// <summary>
        /// Renames the title column
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="spList"></param>
        internal void RenameTitle(ClientContext cContext, SharePointList spList)
        {
            try
            {
                LogVerbose("Start RenameField: " + spList.Name);
                Web wWeb = cContext.Web;
                List lList = wWeb.Lists.GetByTitle(spList.Name);
                FieldCollection collField = lList.Fields;
                Field oneField = collField.GetByInternalNameOrTitle("Title");
                oneField.Title = spList.TitleName;
                oneField.Update();
                cContext.ExecuteQuery();
                LogVerbose("Complete RenameField: " + spList.Name);
            }
            catch (Exception ex)
            {
                LogError(ex, ex.Message);
            }
        }

        /// <summary>
        /// Returns the parent item ID
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="ItemName"></param>
        /// <param name="ParentListName"></param>
        /// <returns></returns>
        internal int GetParentItemID(ClientContext cContext, string ItemName, string ParentListName)
        {
            int nReturn = -1;

            try
            {
                LogVerbose("Start GetParentItemID");

                Web wWeb = cContext.Web;

                List lParentList = cContext.Web.Lists.GetByTitle(ParentListName);

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<Where><Eq><FieldRef Name=\"Title\" /><Value Type=\"Text\">" + ItemName + "</Value></Eq></Where>";
                ListItemCollection collListItem = lParentList.GetItems(camlQuery);

                cContext.Load(collListItem);

                cContext.ExecuteQuery();

                foreach (ListItem oListItem in collListItem)
                {
                    if (oListItem["Title"].ToString() == ItemName)
                    {
                        nReturn = oListItem.Id;
                        break;
                    }
                }

                LogVerbose("Complete GetParentItemID");
            }
            catch (Exception ex)
            {
                LogError(ex, ex.Message);
            }

            return nReturn;
        }

        /// <summary>
        /// Determines if the column is a lookup
        /// </summary>
        /// <param name="List"></param>
        /// <param name="Column"></param>
        /// <param name="Parent"></param>
        /// <returns></returns>
        internal bool IsLookup(SharePointList List, string Column, out string Parent)
        {
            Parent = string.Empty;
            foreach (SharePointColumn spc in List.SharePointColumns)
                if (spc.Name == Column && spc.Parent != string.Empty && spc.Parent != null)
                {
                    Parent = spc.Parent;
                    return true;
                }

            return false;
        }

        /// <summary>
        /// Adds reference data
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="spList"></param>
        internal void AddReferenceData(ClientContext cContext, SharePointList spList)
        {
            try
            {
                LogVerbose("AddReferenceData for: " + spList.Name);

                List lList = cContext.Web.Lists.GetByTitle(spList.Name);
                string strParent = string.Empty;

                foreach (SharePointReferenceRow spRefRow in spList.SharePointReferenceRows)
                {
                    ListItemCreationInformation lici = new ListItemCreationInformation();
                    ListItem newItem = lList.AddItem(lici);

                    foreach (SharePointReferenceColumn spRefCol in spRefRow.SharePointReferenceColumns)
                    {
                        if (IsLookup(spList, spRefCol.Name, out strParent))
                            newItem[spRefCol.Name] = GetParentItemID(cContext, spRefCol.Value, strParent);
                        else
                            newItem[spRefCol.Name] = spRefCol.Value;

                        newItem.Update();
                    }

                    cContext.ExecuteQuery();
                }

                LogVerbose("Completed AddReferenceData for: " + spList.Name);
            }
            catch (Exception ex)
            {
                LogError(ex, ex.Message);
            }
        }

        internal void GenerateCSharp(SharePointContents spContents)
        {
            try
            {
                LogWarning($"Generate C# to {CSharpFile} disk");

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("namespace " + spContents.ApplicationName + "Entities");
                sb.AppendLine("{");
                sb.AppendLine("\t//Autogen'ed from Data on: " + DateTime.Now.ToString());
                sb.AppendLine();

                foreach (SharePointList spList in spContents.SharePointLists)
                {
                    sb.AppendLine("\tpublic class " + spContents.ApplicationName + "Entities_" + spList.Name);
                    sb.AppendLine("\t{");

                    sb.AppendLine("\t\tpublic const string ListName = \"" + spList.Name + "\";");
                    sb.AppendLine();

                    foreach (SharePointColumn spColumn in spList.SharePointColumns)
                    {
                        sb.AppendLine("\t\tpublic const string " + spColumn.Name + " = \"" + spColumn.Name + "\";");
                    }

                    sb.AppendLine("\t\tpublic const string " + spList.TitleName + " = \"Title\";");

                    sb.AppendLine("\t}");
                }
                sb.AppendLine("}");

                using (StreamWriter swFile = new StreamWriter($"{CSharpFile}_Entities.cs", false))
                {
                    swFile.Write(sb.ToString());
                }

                LogVerbose("Completed generate C#");
            }
            catch (Exception ex)
            {
                LogError(ex, ex.Message);
            }
        }

        /// <summary>
        /// Adds a column to a SharePoint list
        /// </summary>
        /// <param name="cContext">Context to SharePoint instance</param>
        /// <param name="ColumnName"></param>
        /// <param name="Type"></param>
        /// <param name="ListName"></param>
        /// <param name="Parent">Optional, this is if you want a reference to another table</param>
        /// <param name="Data">Optional, this is for a choice field</param>
        internal void AddColumn(ClientContext cContext, string ColumnName, string Type, string ListName, string Parent, string Data)
        {
            try
            {
                LogVerbose("Add column: " + ColumnName + " for: " + ListName);
                Web wWeb = cContext.Web;

                string strParentListID = string.Empty;
                if (Parent != null && Parent != string.Empty)
                {
                    List lParentList = cContext.Web.Lists.GetByTitle(Parent);
                    cContext.Load(lParentList);
                    cContext.ExecuteQuery();
                    strParentListID = lParentList.Id.ToString();
                }

                List lList = cContext.Web.Lists.GetByTitle(ListName);
                lList.Description = ListName;

                string strSchema = "<Field DisplayName='" + ColumnName + "' Required='FALSE' Type='" + Type + "' />";

                if (Type == "DateTime")
                    strSchema = "<Field DisplayName='" + ColumnName + "' Format='DateOnly' Required='FALSE' Type='" + Type + "' />";

                if (Type == "Choice")
                {
                    string strChoice = string.Empty;
                    foreach (string strItem in Data.Split(','))
                        strChoice += "<CHOICE>" + strItem + "</CHOICE>";

                    strSchema = @"<Field Type='Choice' DisplayName='" + ColumnName + "' Format='Dropdown'><CHOICES>" + strChoice + "</CHOICES></Field>";
                }

                if (Parent != null && Parent != string.Empty)
                {
                    strSchema = "<Field Type='Lookup' DisplayName='" + ColumnName + "' Required='FALSE' EnforceUniqueValues='FALSE' List='{" + strParentListID + "}' ShowField='Title' UnlimitedLengthInDocumentLibrary='FALSE' RelationshipDeleteBehavior='None' StaticName='" + ColumnName + "' Name='" + ColumnName + "'/>";
                }

                Field fRegion = lList.Fields.AddFieldAsXml(strSchema, true, AddFieldOptions.DefaultValue);
                lList.Update();
                cContext.ExecuteQuery();

                LogVerbose("Complete add column: " + ColumnName + " for: " + ListName);
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("A duplicate field name"))
                {
                    LogError(ex, ex.Message);
                }
            }
        }

        /// <summary>
        /// Deletes a SharePoint list
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="ListName"></param>
        internal void DeleteList(ClientContext cContext, string ListName)
        {
            try
            {
                LogVerbose("Delete list: " + ListName);
                Web wWeb = cContext.Web;
                List lList = wWeb.Lists.GetByTitle(ListName);
                lList.DeleteObject();
                cContext.ExecuteQuery();
                LogVerbose("Delete completed: " + ListName);
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("does not exist at site with URL"))
                    throw (ex);
            }
        }

        /// <summary>
        /// Creates a SharePoint list. Currently it is only supporting 3 types,
        /// feel free to extend
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="ListName"></param>
        /// <param name="ListType"></param>
        internal void CreateList(ClientContext cContext, string ListName, string ListType)
        {
            try
            {
                LogVerbose("Create " + ListName);
                Web wWeb = cContext.Web;

                int nListType = (int)ListTemplateType.GenericList;

                if (ListType == "DocLib")
                    nListType = (int)ListTemplateType.DocumentLibrary;

                if (ListType == "Tasks")
                    nListType = (int)ListTemplateType.Tasks;

                ListCreationInformation lciList = new ListCreationInformation();
                lciList.Title = ListName;
                lciList.TemplateType = nListType;

                List lList = wWeb.Lists.Add(lciList);
                lList.Description = ListName;

                lList.Update();
                cContext.ExecuteQuery();
                LogVerbose("Complete " + ListName);
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("A duplicate field name"))
                {
                    LogError(ex, ex.Message);
                }
            }
        }
    }

    public class SharePointColumn
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Parent { get; set; }
        public string Data { get; set; }
    }

    public class SharePointList
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public bool Ignore { get; set; }
        public string TitleName { get; set; }
        public SharePointColumn[] SharePointColumns { get; set; }
        public SharePointReferenceRow[] SharePointReferenceRows { get; set; }
    }

    public class SharePointReferenceRow
    {
        public SharePointReferenceColumn[] SharePointReferenceColumns { get; set; }
    }
    public class SharePointReferenceColumn
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class SharePointContents
    {
        public string ApplicationName { get; set; }
        public string Server { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
        public bool Iso365 { get; set; }
        public SharePointList[] SharePointLists { get; set; }
    }
}