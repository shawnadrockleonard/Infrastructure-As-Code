using System;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using Microsoft.SharePoint.Client.Utilities;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Core.Extensions;
    using InfrastructureAsCode.Core.Constants;
    using OfficeDevPnP.Core.Utilities;

    /// <summary>
    /// CmdLet will provide a sample query to build HTML based on the view
    /// </summary>
    [Cmdlet(VerbsCommon.Format, "IaCListItemsAsHTML", SupportsShouldProcess = false)]
    public class FormatIaCListItemsAsHTML : IaCCmdlet
    {
        /// <summary>
        /// List Identity
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// View Identity
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public ViewPipeBind View { get; set; }

        /// <summary>
        /// Array of email strings
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public IEnumerable<string> SendTo { get; set; }



        #region Private Region

        /// <summary>
        /// container for column mapping where fieldtype inferrence is required
        /// </summary>
        private List<FieldMappings> ColumnMappings { get; set; }

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var paramlistname = List.Title;
            var paramviewname = View.Title;
            this.ColumnMappings = new List<FieldMappings>();


            try
            {
                var viewlist = ClientContext.Web.GetListByTitle(paramlistname);
                ClientContext.Load(viewlist, rcll => rcll.Fields, rcll => rcll.ItemCount, rcll => rcll.ContentTypes, rcll => rcll.BaseType);
                ClientContext.Load(viewlist.Views, wv => wv.Include(wvi => wvi.Title, wvi => wvi.Id, wvi => wvi.ListViewXml, wvi => wvi.ViewFields));
                ClientContext.ExecuteQueryRetry();

                var viewFieldHeaderHtml = string.Empty;
                var view = viewlist.Views.FirstOrDefault(w => w.Title.Trim().Equals(paramviewname, StringComparison.CurrentCultureIgnoreCase));
                if (view != null)
                {
                    var doc = XDocument.Parse(view.ListViewXml);

                    var queryXml = doc.Root.Element(XName.Get("Query"));
                    var camlFieldXml = doc.Root.Element(XName.Get("ViewFields"));
                    var queryWhereXml = queryXml.Element(XName.Get("Where"));
                    var queryGroupByXml = queryXml.Element(XName.Get("GroupBy"));
                    var queryOrderXml = queryXml.Element(XName.Get("OrderBy"));

                    var queryViewCaml = ((camlFieldXml != null) ? camlFieldXml.ToString() : string.Empty);
                    var queryWhereCaml = ((queryWhereXml != null) ? queryWhereXml.ToString() : string.Empty);
                    var queryOrderCaml = ((queryOrderXml != null) ? queryOrderXml.ToString() : string.Empty);
                    var viewFields = new List<string>() { "ContentTypeId", "FileRef", "FileDirRef", "FileLeafRef" };
                    if (viewlist.BaseType == BaseType.GenericList)
                    {
                        viewFields.AddRange(new string[] { ConstantsListFields.Field_LinkTitle, ConstantsListFields.Field_LinkTitleNoMenu });
                    }
                    if (viewlist.BaseType == BaseType.DocumentLibrary)
                    {
                        viewFields.AddRange(new string[] { ConstantsLibraryFields.Field_LinkFilename, ConstantsLibraryFields.Field_LinkFilenameNoMenu });
                    }
                    foreach (var xnode in camlFieldXml.Descendants())
                    {
                        var attributeValue = xnode.Attribute(XName.Get("Name"));
                        var fe = attributeValue.Value;
                        if (fe == "ContentType")
                        {
                            fe = "ContentTypeId";
                        }

                        if (!viewFields.Any(vf => vf == fe))
                        {
                            viewFields.Add(fe);
                        }
                    }
                    // lets override the view field XML with some additional columns
                    queryViewCaml = CAML.ViewFields(viewFields.Select(s => CAML.FieldRef(s)).ToArray());


                    var viewFieldsHeader = "<tr>";
                    var viewFieldsHeaderIdx = 0;

                    view.ViewFields.ToList().ForEach(fe =>
                    {
                        var fieldDisplayName = viewlist.Fields.FirstOrDefault(fod => fod.InternalName == fe);

                        ColumnMappings.Add(new FieldMappings()
                        {
                            ColumnInternalName = fieldDisplayName.InternalName,
                            ColumnMandatory = fieldDisplayName.Required,
                            ColumnType = fieldDisplayName.FieldTypeKind
                        });

                        viewFieldsHeader += string.Format("<th>{0}</th>", (fieldDisplayName == null ? fe : fieldDisplayName.Title));
                        viewFieldsHeaderIdx++;
                    });

                    viewFieldsHeader += "</tr>";


                    var innerGroupName = string.Empty;
                    var hasGroupStrategy = false;

                    if (queryGroupByXml != null && queryGroupByXml.HasElements)
                    {
                        queryWhereCaml = queryGroupByXml.ToString() + queryWhereCaml;
                        hasGroupStrategy = true;
                        var innerGroupBy = queryGroupByXml.Elements();
                        var innerGroupField = innerGroupBy.FirstOrDefault();
                        innerGroupName = innerGroupField.Attribute(XName.Get("Name")).Value;
                    }

                    var camlQueryXml = CAML.ViewQuery(ViewScope.RecursiveAll, queryWhereCaml, queryOrderCaml, queryViewCaml, 500);

                    ListItemCollectionPosition camlListItemCollectionPosition = null;
                    var camlQuery = new CamlQuery();
                    camlQuery.ViewXml = camlQueryXml;

                    var previousgroupname = "zzzzzzzzzheader";
                    var htmltoemail = new StringBuilder();
                    htmltoemail.Append("<table>");
                    if (!hasGroupStrategy)
                    {
                        htmltoemail.Append(viewFieldsHeader);
                    }

                    while (true)
                    {
                        camlQuery.ListItemCollectionPosition = camlListItemCollectionPosition;
                        var spListItems = viewlist.GetItems(camlQuery);
                        this.ClientContext.Load(spListItems, lti => lti.ListItemCollectionPosition);
                        this.ClientContext.ExecuteQueryRetry();
                        camlListItemCollectionPosition = spListItems.ListItemCollectionPosition;

                        foreach (var ittpItem in spListItems)
                        {
                            LogVerbose("Item {0}", ittpItem.Id);
                            if (hasGroupStrategy)
                            {
                                var currentgroupname = ittpItem.RetrieveListItemValue(innerGroupName).Trim();
                                if (previousgroupname != currentgroupname)
                                {
                                    htmltoemail.AppendFormat("<tr><th colspan='{0}' style='text-align:center;background-color:blue;color:white'>{1}</th></tr>", viewFieldsHeaderIdx, currentgroupname);
                                    htmltoemail.Append(viewFieldsHeader);
                                    previousgroupname = currentgroupname;
                                }
                            }

                            var htmlrow = "<tr>";
                            view.ViewFields.ToList().ForEach(fe =>
                            {
                                if (fe == "ContentType")
                                {
                                    fe = "ContentTypeId";
                                }

                                var htmlrowvalue = string.Empty;
                                try
                                {
                                    var c = ColumnMappings.FirstOrDefault(f => f.ColumnInternalName == fe);
                                    if (c != null && c.ColumnType == FieldType.Lookup)
                                    {
                                        var res = ittpItem.RetrieveListItemValueAsLookup(fe);
                                        htmlrowvalue = res.ToLookupValue();
                                    }
                                    else if (c != null && c.ColumnType == FieldType.User)
                                    {
                                        var res = ittpItem.RetrieveListItemUserValue(fe);
                                        htmlrowvalue = res.ToUserValue();
                                    }
                                    else
                                    {
                                        htmlrowvalue = ittpItem.RetrieveListItemValue(fe);
                                    }
                                }
                                catch (Exception fex) {
                                    LogWarning("Failed to retreive {0} msg => {1}", fe, fex.Message); }
                                finally { }

                                htmlrow += string.Format("<td>{0}</td>", htmlrowvalue);
                            });
                            htmlrow += "</tr>";
                            htmltoemail.Append(htmlrow);
                        }

                        if (camlListItemCollectionPosition == null)
                        {
                            break;
                        }
                    }


                    htmltoemail.Append("</table>");



                    var properties = new EmailProperties
                    {
                        To = SendTo,
                        Subject = $"HTML from Email List ${List.Title}",
                        Body = string.Format("<div>{0}</div>", htmltoemail.ToString())
                    };

                    Microsoft.SharePoint.Client.Utilities.Utility.SendEmail(this.ClientContext, properties);
                    this.ClientContext.ExecuteQueryRetry();
                }

            }
            catch (Exception fex)
            {
                LogError(fex, "Failed to parse view and produce HTML report");
            }

        }
    }
}
