using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Returns the library/list with the column and data
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCFieldColumnAndData")]
    public class GetIaCFieldColumnAndData : IaCAdminCmdlet
    {
        /// <summary>
        /// The Absolute URL for the site we will scan
        /// </summary>
        [Parameter(Mandatory = false)]
        public string SiteUrl { get; set; }


        #region Private Variables

        private string TenantUrl { get; set; }

        private readonly string FieldColumnName = "Name";

        private List<ExtendSPOSiteModel> _siteActionLog { get; set; }

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            _siteActionLog = new List<ExtendSPOSiteModel>();

            try
            {
                TenantContext.EnsureProperties(tssp => tssp.RootSiteUrl);
                TenantUrl = TenantContext.RootSiteUrl;
                LogVerbose("Enumerating Site collections #|# UserID: {0} ......... ", CurrentUserName);

                var collectionOfSites = GetSiteCollections(true);
                foreach (var siteCollection in collectionOfSites.Where(w =>
                    (!string.IsNullOrEmpty(SiteUrl) && w.Url.IndexOf(SiteUrl) > -1) || (string.IsNullOrEmpty(SiteUrl) && 1 == 1)))
                {
                    var _siteUrl = siteCollection.Url;
                    var _totalWebs = siteCollection.WebsCount;
                    LogVerbose("Processing {0} owner {1}", siteCollection.title, siteCollection.Owner);

                    try
                    {
                        SetSiteAdmin(_siteUrl, CurrentUserName, true);

                        using (var siteContext = this.ClientContext.Clone(_siteUrl))
                        {
                            Web _web = siteContext.Web;
                            var extendedModel = new ExtendSPOSiteModel(siteCollection);

                            extendedModel = ProcessSiteCollectionSubWeb(extendedModel, _web, true);

                            // Add the Site collection to the report
                            _siteActionLog.Add(extendedModel);
                        }
                    }
                    catch (Exception e)
                    {
                        LogError(e, "Failed to processSiteCollection with url {0}", _siteUrl);
                    }
                    finally
                    {
                        //SetSiteAdmin(_siteUrl, CurrentUserName);
                    }
                }

                WriteObject(_siteActionLog, true);
            }
            catch (Exception e)
            {
                LogError(e, "Failed in SetEveryoneGroup cmdlet {0}", e.Message);
            }
        }

        /// <summary>
        /// Process the site subweb
        /// </summary>
        /// <param name="model"></param>
        /// <param name="_web"></param>
        /// <param name="isSiteCollection">(OPTIONAL) indicates the top level</param>
        private dynamic ProcessSiteCollectionSubWeb(dynamic model, Web _web, bool isSiteCollection = false)
        {
            try
            {
                _web.EnsureProperties(spp => spp.Id, spp => spp.Url);
                var _siteUrl = _web.Url;
                var _rootSiteIndex = _siteUrl.ToLower().IndexOf(TenantUrl);

                LogVerbose("Processing web URL {0} SKIPPING:{1}", _siteUrl, (_rootSiteIndex == -1));
                if (_rootSiteIndex > -1)
                {

                    var sitemodel = ProcessSite(_web);
                    if (sitemodel.Id.HasValue)
                    {
                        model.Sites.Add(sitemodel);
                    }

                    //Process subsites
                    _web.Context.Load(_web.Webs);
                    _web.Context.ExecuteQueryRetry();

                    if (model.WebsCount > 1)
                    {
                        LogVerbose("Site {0} has webs {1}", _siteUrl, model.WebsCount);
                        foreach (Web _inWeb in _web.Webs)
                        {
                            model = ProcessSiteCollectionSubWeb(model, _inWeb);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e, "Failed in processSiteCollection");
            }

            return model;
        }


        private SPSiteModel ProcessSite(Web _web)
        {
            var hasListFound = false;
            var model = new SPSiteModel();

            _web.EnsureProperties(wssp => wssp.Id,
                wspp => wspp.ServerRelativeUrl,
                wspp => wspp.Title,
                wssp => wssp.HasUniqueRoleAssignments,
                wssp => wssp.SiteUsers,
                wssp => wssp.Url,
                wssp => wssp.Lists,
                wssp => wssp.ContentTypes.Include(
                    lcnt => lcnt.Id,
                    lcnt => lcnt.Name,
                    lcnt => lcnt.StringId,
                    lcnt => lcnt.Description,
                    lcnt => lcnt.DocumentTemplate,
                    lcnt => lcnt.Group,
                    lcnt => lcnt.Hidden,
                    lcnt => lcnt.JSLink,
                    lcnt => lcnt.SchemaXml,
                    lcnt => lcnt.Scope,
                    lcnt => lcnt.FieldLinks.Include(
                        lcntlnk => lcntlnk.Id,
                        lcntlnk => lcntlnk.Name,
                        lcntlnk => lcntlnk.Hidden,
                        lcntlnk => lcntlnk.Required
                        ),
                    lcnt => lcnt.Fields.Include(
                        lcntfld => lcntfld.FieldTypeKind,
                        lcntfld => lcntfld.InternalName,
                        lcntfld => lcntfld.Id,
                        lcntfld => lcntfld.Group,
                        lcntfld => lcntfld.Title,
                        lcntfld => lcntfld.Hidden,
                        lcntfld => lcntfld.Description,
                        lcntfld => lcntfld.JSLink,
                        lcntfld => lcntfld.Indexed,
                        lcntfld => lcntfld.Required,
                        lcntfld => lcntfld.SchemaXml)),
                wssp => wssp.Fields.Include(
                    lcntfld => lcntfld.FieldTypeKind,
                    lcntfld => lcntfld.InternalName,
                    lcntfld => lcntfld.Id,
                    lcntfld => lcntfld.Group,
                    lcntfld => lcntfld.Title,
                    lcntfld => lcntfld.Hidden,
                    lcntfld => lcntfld.Description,
                    lcntfld => lcntfld.JSLink,
                    lcntfld => lcntfld.Indexed,
                    lcntfld => lcntfld.Required,
                    lcntfld => lcntfld.SchemaXml));
            model.Url = _web.Url;
            model.title = _web.Title;
            LogVerbose("Processing: {0}", _web.Url);


            /* Process Fields */
            try
            {
                foreach (var _fields in _web.Fields)
                {
                    if (string.IsNullOrEmpty(FieldColumnName) ||
                        _fields.Title.Equals(FieldColumnName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        hasListFound = true;

                        model.FieldDefinitions.Add(new SPFieldDefinitionModel()
                        {
                            FieldTypeKind = _fields.FieldTypeKind,
                            InternalName = _fields.InternalName,
                            FieldGuid = _fields.Id,
                            GroupName = _fields.Group,
                            Title = _fields.Title,
                            HiddenField = _fields.Hidden,
                            Description = _fields.Description,
                            JSLink = _fields.JSLink,
                            FieldIndexed = _fields.Indexed,
                            Required = _fields.Required,
                            SchemaXml = _fields.SchemaXml
                        });
                    }
                };
            }
            catch (Exception e)
            {
                LogError(e, "Failed to retrieve site owners {0}", _web.Url);
            }

            /* Process Content Type */
            try
            {
                foreach (var _ctypes in _web.ContentTypes)
                {
                    var cmodel = new SPContentTypeDefinition()
                    {
                        ContentTypeId = _ctypes.StringId,
                        Name = _ctypes.Name,
                        Description = _ctypes.Description,
                        DocumentTemplate = _ctypes.DocumentTemplate,
                        ContentTypeGroup = _ctypes.Group,
                        Hidden = _ctypes.Hidden,
                        JSLink = _ctypes.JSLink,
                        Scope = _ctypes.Scope
                    };

                    foreach (var _ctypeFields in _ctypes.FieldLinks)
                    {
                        if (string.IsNullOrEmpty(FieldColumnName) ||
                            _ctypeFields.Name.Equals(FieldColumnName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            cmodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                            {
                                Name = _ctypeFields.Name,
                                Id = _ctypeFields.Id,
                                Hidden = _ctypeFields.Hidden,
                                Required = _ctypeFields.Required
                            });
                        }
                    }

                    if (cmodel.FieldLinks.Any())
                    {
                        hasListFound = true;
                        model.ContentTypes.Add(cmodel);
                    }
                };
            }
            catch (Exception e)
            {
                LogError(e, "Failed to retrieve site owners {0}", _web.Url);
            }

            // ********** Process List
            try
            {
                var lists = ProcessList(_web);
                if (lists.Any())
                {
                    hasListFound = true;
                    model.Lists.AddRange(lists);
                }
            }
            catch (Exception e)
            {
                LogError(e, "Exception occurred in processSite");
            }

            if (hasListFound)
            {
                // setting ID to indicate to parent consumer that this entity has unique permissions in the TREE
                model.Id = _web.Id;
            }

            return model;
        }

        private IList<SPListDefinition> ProcessList(Web _web)
        {
            var model = new List<SPListDefinition>();

            // ********** Process Lists
            ListCollection _lists = _web.Lists;
            _web.Context.Load(_lists,
                spp => spp.Include(
                    sppi => sppi.Id,
                    sppi => sppi.Title,
                    sppi => sppi.RootFolder.ServerRelativeUrl,
                    sppi => sppi.HasUniqueRoleAssignments,
                    sppi => sppi.BaseTemplate,
                    sppi => sppi.Hidden,
                    sppi => sppi.IsSystemList,
                    sppi => sppi.IsPrivate,
                    sppi => sppi.IsApplicationList,
                    sppi => sppi.IsCatalog,
                    sppi => sppi.IsSiteAssetsLibrary));

            _web.Context.ExecuteQueryRetry();

            var docLibEnumValue = Convert.ToInt32(Microsoft.SharePoint.Client.ListTemplateType.DocumentLibrary);

            // Restrict to natural lists or custom lists
            foreach (List _list in _lists.Where(sppi
                            => !sppi.IsSystemList
                            && !sppi.IsApplicationList
                            && sppi.BaseTemplate != docLibEnumValue))
            {
                var hasListFound = false;
                var listContext = _list.Context;
                LogVerbose("Enumerating List {0} URL:{1}", _list.Title, _list.RootFolder.ServerRelativeUrl);

                try
                {
                    listContext.Load(_list,
                        lssp => lssp.Id,
                        lssp => lssp.Title,
                        lssp => lssp.HasUniqueRoleAssignments,
                        lssp => lssp.Title,
                        lssp => lssp.Hidden,
                        lssp => lssp.IsSystemList,
                        lssp => lssp.IsPrivate,
                        lssp => lssp.IsApplicationList,
                        lssp => lssp.IsCatalog,
                        lssp => lssp.IsSiteAssetsLibrary,
                        lssp => lssp.RootFolder.ServerRelativeUrl,
                        lssp => lssp.ContentTypes.Include(
                            lcnt => lcnt.Id,
                            lcnt => lcnt.Name,
                            lcnt => lcnt.StringId,
                            lcnt => lcnt.Description,
                            lcnt => lcnt.DocumentTemplate,
                            lcnt => lcnt.Group,
                            lcnt => lcnt.Hidden,
                            lcnt => lcnt.JSLink,
                            lcnt => lcnt.SchemaXml,
                            lcnt => lcnt.Scope,
                            lcnt => lcnt.FieldLinks.Include(
                                lcntlnk => lcntlnk.Id,
                                lcntlnk => lcntlnk.Name,
                                lcntlnk => lcntlnk.Hidden,
                                lcntlnk => lcntlnk.Required
                                ),
                            lcnt => lcnt.Fields.Include(
                                lcntfld => lcntfld.FieldTypeKind,
                                lcntfld => lcntfld.InternalName,
                                lcntfld => lcntfld.Id,
                                lcntfld => lcntfld.Group,
                                lcntfld => lcntfld.Title,
                                lcntfld => lcntfld.Hidden,
                                lcntfld => lcntfld.Description,
                                lcntfld => lcntfld.JSLink,
                                lcntfld => lcntfld.Indexed,
                                lcntfld => lcntfld.Required,
                                lcntfld => lcntfld.SchemaXml)),
                        lssp => lssp.Fields.Include(
                            lcntfld => lcntfld.FieldTypeKind,
                            lcntfld => lcntfld.InternalName,
                            lcntfld => lcntfld.Id,
                            lcntfld => lcntfld.Group,
                            lcntfld => lcntfld.Title,
                            lcntfld => lcntfld.Hidden,
                            lcntfld => lcntfld.Description,
                            lcntfld => lcntfld.JSLink,
                            lcntfld => lcntfld.Indexed,
                            lcntfld => lcntfld.Required,
                            lcntfld => lcntfld.SchemaXml));
                    listContext.ExecuteQueryRetry();

                    var listModel = new SPListDefinition()
                    {
                        Id = _list.Id,
                        HasUniquePermission = _list.HasUniqueRoleAssignments,
                        ListName = _list.Title,
                        ServerRelativeUrl = _list.RootFolder.ServerRelativeUrl,
                        Hidden = _list.Hidden,
                        IsSystemList = _list.IsSystemList,
                        IsPrivate = _list.IsPrivate,
                        IsApplicationList = _list.IsApplicationList,
                        IsCatalog = _list.IsCatalog,
                        IsSiteAssetsLibrary = _list.IsSiteAssetsLibrary
                    };

                    /* Process Fields */
                    try
                    {
                        foreach (var _fields in _list.Fields)
                        {
                            if (string.IsNullOrEmpty(FieldColumnName) ||
                                _fields.Title.Equals(FieldColumnName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                hasListFound = true;

                                listModel.FieldDefinitions.Add(new SPFieldDefinitionModel()
                                {
                                    FieldTypeKind = _fields.FieldTypeKind,
                                    InternalName = _fields.InternalName,
                                    FieldGuid = _fields.Id,
                                    GroupName = _fields.Group,
                                    Title = _fields.Title,
                                    HiddenField = _fields.Hidden,
                                    Description = _fields.Description,
                                    JSLink = _fields.JSLink,
                                    FieldIndexed = _fields.Indexed,
                                    Required = _fields.Required,
                                    SchemaXml = _fields.SchemaXml
                                });
                            }
                        };
                    }
                    catch (Exception e)
                    {
                        LogError(e, "Failed to retrieve site owners {0}", _web.Url);
                    }

                    foreach (var _ctypes in _list.ContentTypes)
                    {
                        var cmodel = new SPContentTypeDefinition()
                        {
                            ContentTypeId = _ctypes.StringId,
                            Name = _ctypes.Name,
                            Description = _ctypes.Description,
                            DocumentTemplate = _ctypes.DocumentTemplate,
                            ContentTypeGroup = _ctypes.Group,
                            Hidden = _ctypes.Hidden,
                            JSLink = _ctypes.JSLink,
                            Scope = _ctypes.Scope
                        };

                        foreach (var _ctypeFields in _ctypes.FieldLinks)
                        {
                            if (string.IsNullOrEmpty(FieldColumnName) ||
                                _ctypeFields.Name.Equals(FieldColumnName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                cmodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                {
                                    Name = _ctypeFields.Name,
                                    Id = _ctypeFields.Id,
                                    Hidden = _ctypeFields.Hidden,
                                    Required = _ctypeFields.Required
                                });
                            }
                        }

                        if (cmodel.FieldLinks.Any())
                        {
                            hasListFound = true;
                            listModel.ContentTypes.Add(cmodel);
                        }
                    }

                    if (hasListFound)
                    {
                        model.Add(listModel);
                    }
                }
                catch (Exception e)
                {
                    LogError(e, "Failed in ProcessList");
                }
            }
            return model;
        }
    }
}
