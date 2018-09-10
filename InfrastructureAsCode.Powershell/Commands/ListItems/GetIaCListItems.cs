using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;

    [Cmdlet(VerbsCommon.Get, "IaCListItems")]
    public class GetIaCListItems : IaCWebCmdlet
    {
        private const string ParameterSet_BYID = "By Id";
        private const string ParameterSet_BYUNIQUEID = "By Unique Id";
        private const string ParameterSet_BYQUERY = "By Query";
        private const string ParameterSet_ALLITEMS = "All Items";
        [Parameter(Mandatory = true, HelpMessage = "The list to query", Position = 0, ParameterSetName = ParameterAttribute.AllParameterSets)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The ID of the item to retrieve", ParameterSetName = ParameterSet_BYID)]
        public int Id = -1;

        [Parameter(Mandatory = false, HelpMessage = "The unique id (GUID) of the item to retrieve", ParameterSetName = ParameterSet_BYUNIQUEID)]
        public GuidPipeBind UniqueId { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The CAML query to execute against the list", ParameterSetName = ParameterSet_BYQUERY)]
        public string Query { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_ALLITEMS)]
        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_BYID)]
        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_BYUNIQUEID)]
        public string[] Fields { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The number of items to retrieve per page request.", ParameterSetName = ParameterSet_ALLITEMS)]
        [Parameter(Mandatory = false, HelpMessage = "The number of items to retrieve per page request.", ParameterSetName = ParameterSet_BYQUERY)]
        public int PageSize = -1;

        [Parameter(Mandatory = false, HelpMessage = "The script block to run after every page request.", ParameterSetName = ParameterSet_ALLITEMS)]
        [Parameter(Mandatory = false, HelpMessage = "The script block to run after every page request.", ParameterSetName = ParameterSet_BYQUERY)]
        public ScriptBlock ScriptBlock { get; set; }


        public override void ExecuteCmdlet()
        {
            var list = List.GetList(SelectedWeb);

            if (HasId())
            {
                var listItem = list.GetItemById(Id);
                if (Fields != null)
                {
                    foreach (var field in Fields)
                    {
                        ClientContext.Load(listItem, l => l[field]);
                    }
                }
                else
                {
                    ClientContext.Load(listItem);
                }
                ClientContext.ExecuteQueryRetry();
                WriteObject(listItem);
            }
            else if (HasUniqueId())
            {
                CamlQuery query = new CamlQuery();
                var viewFieldsStringBuilder = new StringBuilder();
                if (HasFields())
                {
                    viewFieldsStringBuilder.Append("<ViewFields>");
                    foreach (var field in Fields)
                    {
                        viewFieldsStringBuilder.AppendFormat("<FieldRef Name='{0}'/>", field);
                    }
                    viewFieldsStringBuilder.Append("</ViewFields>");
                }
                query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='GUID'/><Value Type='Guid'>{UniqueId.Id}</Value></Eq></Where></Query>{viewFieldsStringBuilder}</View>";
                var listItem = list.GetItems(query);
                ClientContext.Load(listItem);
                ClientContext.ExecuteQueryRetry();
                WriteObject(listItem);
            }
            else
            {
                CamlQuery query = HasCamlQuery() ? new CamlQuery { ViewXml = Query } : CamlQuery.CreateAllItemsQuery();

                if (Fields != null)
                {
                    var queryElement = XElement.Parse(query.ViewXml);

                    var viewFields = queryElement.Descendants("ViewFields").FirstOrDefault();
                    if (viewFields != null)
                    {
                        viewFields.RemoveAll();
                    }
                    else
                    {
                        viewFields = new XElement("ViewFields");
                        queryElement.Add(viewFields);
                    }

                    foreach (var field in Fields)
                    {
                        XElement viewField = new XElement("FieldRef");
                        viewField.SetAttributeValue("Name", field);
                        viewFields.Add(viewField);
                    }
                    query.ViewXml = queryElement.ToString();
                }

                if (HasPageSize())
                {
                    var queryElement = XElement.Parse(query.ViewXml);

                    var rowLimit = queryElement.Descendants("RowLimit").FirstOrDefault();
                    if (rowLimit != null)
                    {
                        rowLimit.RemoveAll();
                    }
                    else
                    {
                        rowLimit = new XElement("RowLimit");
                        queryElement.Add(rowLimit);
                    }

                    rowLimit.SetAttributeValue("Paged", "TRUE");
                    rowLimit.SetValue(PageSize);

                    query.ViewXml = queryElement.ToString();
                }

                do
                {
                    var listItems = list.GetItems(query);
                    ClientContext.Load(listItems);
                    ClientContext.ExecuteQueryRetry();

                    WriteObject(listItems, true);

                    if (ScriptBlock != null)
                    {
                        ScriptBlock.Invoke(listItems);
                    }

                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null);
            }
        }

        private bool HasId()
        {
            return Id != -1;
        }

        private bool HasUniqueId()
        {
            return UniqueId != null && UniqueId.Id != Guid.Empty;
        }

        private bool HasCamlQuery()
        {
            return Query != null;
        }

        private bool HasFields()
        {
            return Fields != null;
        }

        private bool HasPageSize()
        {
            return PageSize > 0;
        }
    }
}