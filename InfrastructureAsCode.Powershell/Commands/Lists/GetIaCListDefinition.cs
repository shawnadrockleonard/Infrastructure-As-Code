using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Models.Minimal;
using InfrastructureAsCode.Core.Reports;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Returns the list definition, views, columns, settings
    /// </summary>
    /// <remarks>
    /// Get-IaCListDefinition -List ""Demo List""
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListDefinition")]
    [OutputType(typeof(SPListDefinition))]
    public class GetIaCListDefinition : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity;

        /// <summary>
        /// Expand the list definition
        /// </summary>
        [Parameter(Mandatory = false, Position = 1)]
        public SwitchParameter ExpandObjects { get; set; }



        public override void ExecuteCmdlet()
        {

            // Initialize logging instance with Powershell logger
            ITraceLogger logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);


            // Skip these specific fields
            var skiptypes = new FieldType[]
            {
                FieldType.Invalid,
                FieldType.Computed,
                FieldType.ContentTypeId,
                FieldType.Invalid,
                FieldType.WorkflowStatus,
                FieldType.WorkflowEventType,
                FieldType.Threading,
                FieldType.ThreadIndex,
                FieldType.Recurrence,
                FieldType.PageSeparator,
                FieldType.OutcomeChoice,
                FieldType.CrossProjectLink,
                FieldType.ModStat,
                FieldType.Error,
                FieldType.MaxItems,
                FieldType.Attachments
            };

            // Construct the model
            var SiteComponents = new SiteProvisionerModel()
            {
                FieldChoices = new List<SiteProvisionerFieldChoiceModel>(),
                Lists = new List<SPListDefinition>()
            };


            if (Identity != null)
            {
                var list = Identity.GetList(this.ClientContext.Web);
                if (list != null)
                {
                    var _ctx = this.ClientContext;
                    var _contextWeb = this.ClientContext.Web;
                    var _site = this.ClientContext.Site;

                    ClientContext.Load(_contextWeb, ctxw => ctxw.ServerRelativeUrl, ctxw => ctxw.Id);
                    ClientContext.Load(_site, cts => cts.Id);
                    ClientContext.ExecuteQueryRetry();


                    var weburl = TokenHelper.EnsureTrailingSlash(_contextWeb.ServerRelativeUrl);


                    var listmodel = ClientContext.GetListDefinition(_contextWeb, list, ExpandObjects, logger, skiptypes, null);


                    SiteComponents.Lists.Add(listmodel);
                }
            }

            // Write the model to memory
            WriteObject(SiteComponents);
        }
    }
}
