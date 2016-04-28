using Microsoft.IaC.Core.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Microsoft.IaC.Core.Extensions
{
    public static partial class ListExtensions
    {
        /// <summary>
        /// Adds a field to a list
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="fieldAsXml">The XML declaration of SiteColumn definition</param>
        /// <param name="AddToDefaultView">Optionally add to the default view</param>
        /// <param name="executeQuery">Optionally skip the executeQuery action</param>
        /// <returns>The newly created field or existing field.</returns>
        public static Field CreateField(this List list, string fieldAsXml, bool AddToDefaultView = false, bool executeQuery = true)
        {
            var fields = list.Fields;
            list.Context.Load(fields);
            list.Context.ExecuteQueryRetry();

            var xd = XDocument.Parse(fieldAsXml);
            if (xd.Root != null)
            {
                var ns = xd.Root.Name.Namespace;

                var fieldNode = (from f in xd.Elements(ns + "Field") select f).FirstOrDefault();

                if (fieldNode != null)
                {
                    string id = string.Empty;
                    if (fieldNode.Attribute("ID") != null)
                    {
                        id = fieldNode.Attribute("ID").Value;
                    }
                    else
                    {
                        id = "<No ID specified in XML>";
                    }
                    var name = fieldNode.Attribute("Name").Value;

                    Log.Info("FieldAndContentTypeExtensions", "CreateField {0} with ID {1}", name, id);
                }
            }
            var field = fields.AddFieldAsXml(fieldAsXml, AddToDefaultView, AddFieldOptions.AddFieldInternalNameHint);
            list.Update();

            if (executeQuery)
            {
                list.Context.ExecuteQueryRetry();
            }

            return field;
        }


        /// <summary>
        /// Build folder path from root of the parent list
        /// </summary>
        /// <param name="parentList"></param>
        /// <param name="folderUrl"></param>
        /// <returns></returns>
        public static Folder ListEnsureFolder(this List parentList, string folderUrl)
        {
            var folderNames = folderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var folderName = folderNames[0];

            var ctx = parentList.Context;
            if (!parentList.IsPropertyAvailable("RootFolder"))
            {
                ctx.Load(parentList.RootFolder);
                ctx.ExecuteQueryRetry();
            }

            var folder = parentList.RootFolder.ListEnsureFolder(folderUrl);
            return folder;
        }

        /// <summary>
        /// Build folder path
        /// </summary>
        /// <param name="parentFolder"></param>
        /// <param name="folderUrl"></param>
        /// <returns></returns>
        public static Folder ListEnsureFolder(this Folder parentFolder, string folderUrl)
        {
            var folderNames = folderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var folderName = folderNames[0];

            var ctx = parentFolder.Context;
            if (!parentFolder.IsPropertyAvailable("Folders"))
            {
                ctx.Load(parentFolder.Folders);
                ctx.ExecuteQueryRetry();
            }

            var folder = parentFolder.EnsureFolder(folderName);

            if (folderNames.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderNames, 1, folderNames.Length - 1);
                return ListEnsureFolder(folder, subFolderUrl);
            }

            return folder;
        }

        /// <summary>
        /// get the xml for an xslt web part
        /// </summary>
        /// <param name="viewList">ID of the list</param>
        /// <param name="pageUrl">relative page url</param>
        /// <param name="title">title of the list</param>
        /// <param name="viewID">Represents the View base for the webpart</param>
        /// <returns>string</returns>
        public static string GetXsltWebPartXML(this List viewList, string pageUrl, string title, Guid viewID)
        {
            var executor = false;
            if (!viewList.IsPropertyAvailable("Id"))
            {
                viewList.Context.Load(viewList, vl => vl.Id, vl => vl.Title);
                executor = true;
            }
            if (!viewList.IsPropertyAvailable("RootFolder"))
            {
                viewList.Context.Load(viewList.RootFolder, rf => rf.ServerRelativeUrl, rf => rf.ItemCount, rf => rf.Name);
                executor = true;
            }

            // The properties were not loaded from caller
            if (executor)
            {
                viewList.Context.ExecuteQueryRetry();
            }

            Guid listID = viewList.Id;
            var listUrl = viewList.RootFolder.ServerRelativeUrl;


            StringBuilder wp = new StringBuilder(100);
            wp.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            wp.Append("<webParts>");
            wp.Append("	<webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">");
            wp.Append("		<metaData>");
            wp.Append("			<type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />");
            wp.Append("			<importErrorMessage>Cannot import this Web Part.</importErrorMessage>");
            wp.Append("		</metaData>");
            wp.Append("		<data>");
            wp.Append("			<properties>");
            wp.Append("				<property name=\"Default\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"IsIncluded\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"NoDefaultStyle\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"ViewContentTypeId\" type=\"string\" />");
            wp.AppendFormat("		<property name=\"ListUrl\" type=\"string\">{0}</property>", pageUrl);
            wp.AppendFormat("		<property name=\"ListId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">{0}</property>", listID.ToString());
            wp.AppendFormat("		<property name=\"TitleUrl\" type=\"string\">{0}</property>", listUrl);
            wp.AppendFormat("		<property name=\"ListName\" type=\"string\">{0}</property>", listID.ToString("B").ToUpper());
            wp.AppendFormat("		<property name=\"Title\" type=\"string\">{0}</property>", title);
            wp.Append("             <property name=\"Toolbar Type\" type=\"string\">None</property>");
            wp.Append("				<property name=\"PageType\" type=\"Microsoft.SharePoint.PAGETYPE, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">PAGE_NORMALVIEW</property>");
            wp.AppendFormat("       <property name=\"ViewGuid\" type=\"string\">{0}</property>", viewID.ToString("B").ToUpper());
            wp.Append("				<property name=\"XmlDefinition\" type=\"string\">");
            wp.AppendFormat("&lt;View Name=\"{1}\" Type=\"HTML\" Hidden=\"TRUE\" ReadOnly=\"TRUE\" OrderedView=\"TRUE\" DisplayName=\"\" Url=\"{0}\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" &gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=\"Title\" Ascending=\"TRUE\"/&gt;&lt;FieldRef Name=\"Modified\" Ascending=\"FALSE\"/&gt;&lt;/OrderBy&gt;&lt;/Query&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"Title\"/&gt;&lt;/ViewFields&gt;&lt;RowLimit Paged=\"TRUE\"&gt;30&lt;/RowLimit&gt;&lt;JSLink&gt;sp.ui.tileview.js&lt;/JSLink&gt;&lt;XslLink Default=\"TRUE\"&gt;main.xsl&lt;/XslLink&gt;&lt;Toolbar Type=\"None\"/&gt;&lt;/View&gt;",
                pageUrl, Guid.NewGuid().ToString("B").ToUpper());
            wp.Append("             </property>");
            wp.Append("			</properties>");
            wp.Append("		</data>");
            wp.Append("	</webPart>");
            wp.Append("</webParts>");
            return wp.ToString();
        }

        /// <summary>
        /// Provisions a column based on the field defintion to the host list
        /// </summary>
        /// <param name="hostList">The instantiated list/library to which the field will be added</param>
        /// <param name="fieldDef">The definition for the field</param>
        /// <param name="logger">Provides a method for logging</param>
        /// <param name="SiteGroups">(OPTIONAL) collection of group, required if this is a PeoplePicker field</param>
        /// <param name="JsonFilePath">(OPTIONAL) file path except if loading choices from JSON</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">For field definitions that do not contain all required data</exception>
        public static Field CreateListColumn(this List hostList, SPFieldDefinitionModel fieldDef, Action<string, string[]> loggerVerbose, Action<string, string[]> loggerError, List<SPGroupDefinitionModel> SiteGroups, string JsonFilePath = null)
        {
            var idguid = fieldDef.FieldGuid;
            var choiceXml = string.Empty;
            var defaultChoiceXml = string.Empty;
            var attributes = new List<KeyValuePair<string, string>>();

            if (string.IsNullOrEmpty(fieldDef.InternalName))
            {
                throw new ArgumentNullException("InternalName");
            }

            if (string.IsNullOrEmpty(fieldDef.DisplayName))
            {
                throw new ArgumentNullException("DisplayName");
            }

            if(!string.IsNullOrEmpty(fieldDef.LoadFromJSON) && string.IsNullOrEmpty(JsonFilePath))
            {
                throw new ArgumentNullException("JsonFilePath", "You must specify a file path to the JSON file if loading from JSON");
            }

            if(!string.IsNullOrEmpty(fieldDef.PeopleGroupName) && (SiteGroups == null || SiteGroups.Count() <= 0))
            {
                throw new ArgumentNullException("SiteGroups", string.Format("You must specify a collection of group for the field {0}", fieldDef.DisplayName));
            }


            if (!hostList.IsPropertyAvailable("Context"))
            {
                
            }

            var fields = hostList.Fields;
            hostList.Context.Load(fields, fc => fc.Include(f => f.Id, f => f.InternalName, f => f.Title, f => f.JSLink, f => f.Indexed, f => f.CanBeDeleted, f => f.Required));
            hostList.Context.ExecuteQueryRetry();

            var returnField = fields.FirstOrDefault(f => f.Id == fieldDef.FieldGuid || f.InternalName == fieldDef.InternalName);
            if (returnField == null)
            {
                var finfo = fieldDef.ToCreationObject();

                try
                {
                    if (!string.IsNullOrEmpty(fieldDef.Description))
                    {
                        attributes.Add(new KeyValuePair<string, string>("Description", fieldDef.Description));
                    }
                    if (fieldDef.FieldIndexed)
                    {
                        attributes.Add(new KeyValuePair<string, string>("Indexed", fieldDef.FieldIndexed.ToString().ToUpper()));
                    }

                    var choices = new FieldType[] { FieldType.Choice, FieldType.GridChoice, FieldType.MultiChoice, FieldType.OutcomeChoice };
                    if (choices.Any(a => a == fieldDef.fieldType))
                    {
                        if (!string.IsNullOrEmpty(fieldDef.LoadFromJSON))
                        {
                            var filePath = string.Format("{0}\\{1}", JsonFilePath, fieldDef.LoadFromJSON);
                            //#TODO: Check file path
                            var contents = JsonConvert.DeserializeObject<List<SPChoiceModel>>(System.IO.File.ReadAllText(filePath));
                            fieldDef.FieldChoices.Clear();
                            fieldDef.FieldChoices.AddRange(contents);
                        }

                        choiceXml = string.Format("<CHOICES>{0}</CHOICES>", string.Join("", fieldDef.FieldChoices.Select(s => string.Format("<CHOICE>{0}</CHOICE>", s.Choice.Trim())).ToArray()));
                        if (!string.IsNullOrEmpty(fieldDef.ChoiceDefault))
                        {
                            defaultChoiceXml = string.Format("<Default>{0}</Default>", fieldDef.ChoiceDefault);
                        }
                        if (fieldDef.fieldType == FieldType.Choice)
                        {
                            attributes.Add(new KeyValuePair<string, string>("Format", fieldDef.ChoiceFormat.ToString("f")));
                        }

                    }
                    else if (fieldDef.fieldType == FieldType.DateTime)
                    {
                        if (fieldDef.DateFieldFormat.HasValue)
                        {
                            attributes.Add(new KeyValuePair<string, string>("DisplayFormat", fieldDef.DateFieldFormat.Value.ToString("f")));
                        }
                    }
                    else if (fieldDef.fieldType == FieldType.Note)
                    {
                        attributes.Add(new KeyValuePair<string, string>("RichText", fieldDef.RichTextField.ToString().ToUpper()));
                        attributes.Add(new KeyValuePair<string, string>("RestrictedMode", fieldDef.RestrictedMode.ToString().ToUpper()));
                        attributes.Add(new KeyValuePair<string, string>("NumLines", fieldDef.NumLines.ToString()));
                        if (!fieldDef.RestrictedMode)
                        {
                            attributes.Add(new KeyValuePair<string, string>("RichTextMode", "FullHtml"));
                            attributes.Add(new KeyValuePair<string, string>("IsolateStyles", true.ToString().ToUpper()));
                        }

                    }
                    else if (fieldDef.fieldType == FieldType.User)
                    {
                        //AllowMultipleValues
                        if (fieldDef.MultiChoice)
                        {
                            attributes.Add(new KeyValuePair<string, string>("Mult", fieldDef.MultiChoice.ToString().ToUpper()));
                        }
                        //SelectionMode
                        if (fieldDef.PeopleOnly)
                        {
                            attributes.Add(new KeyValuePair<string, string>("UserSelectionMode", FieldUserSelectionMode.PeopleOnly.ToString("d")));
                        }

                        if (!string.IsNullOrEmpty(fieldDef.PeopleLookupField))
                        {
                            attributes.Add(new KeyValuePair<string, string>("ShowField", fieldDef.PeopleLookupField));
                            //fldUser.LookupField = fieldDef.PeopleLookupField;
                        }
                        if (!string.IsNullOrEmpty(fieldDef.PeopleGroupName))
                        {
                            var group = SiteGroups.FirstOrDefault(f => f.Title == fieldDef.PeopleGroupName);
                            if (group != null)
                            {
                                attributes.Add(new KeyValuePair<string, string>("UserSelectionScope", group.Id.ToString()));
                            }
                        }
                    }


                    finfo.AdditionalAttributes = attributes;
                    var finfoXml = FieldAndContentTypeExtensions.FormatFieldXml(finfo);
                    if (!string.IsNullOrEmpty(choiceXml))
                    {
                        XDocument xd = XDocument.Parse(finfoXml);
                        XElement root = xd.FirstNode as XElement;
                        if (!string.IsNullOrEmpty(defaultChoiceXml))
                        {
                            root.Add(XElement.Parse(defaultChoiceXml));
                        }
                        root.Add(XElement.Parse(choiceXml));
                        finfoXml = xd.ToString();
                    }
                    loggerVerbose("Provision field {0} with XML:{1}", new string[] { finfo.InternalName, finfoXml });

                    // Should throw an exception if the field ID or Name exist in the list
                    var baseField = hostList.CreateField(finfoXml, finfo.AddToDefaultView, executeQuery: false);
                    hostList.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    var msg = ex.Message;
                    loggerError("EXCEPTION: field {0} with message {1}", new string[] { fieldDef.InternalName, msg });
                }
                finally
                {
                    returnField = hostList.Fields.GetByInternalNameOrTitle(fieldDef.InternalName);
                    hostList.Context.Load(returnField, fd => fd.Id, fd => fd.Title, fd => fd.Indexed, fd => fd.InternalName, fd => fd.CanBeDeleted, fd => fd.Required);
                    hostList.Context.ExecuteQuery();
                }
            }

            return returnField;
        }
    }
}
