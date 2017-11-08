using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// The function cmdlet will set the JS Link for the field definition
    /// </summary>
    /// <remarks>
    ///     Set-IaCListFieldJsLink -Identity "List Title" -FieldName "Internal_x0020_Name" -JsLink "/SiteAssets/CSS/NewSite.css" 
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCListFieldJsLink", SupportsShouldProcess = true)]
    public class SetIaCListFieldJsLink : IaCCmdlet
    {
        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind Identity { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public FieldPipeBind FieldIdentity { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string JsLink { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var _web = this.ClientContext.Web;
            var _list = Identity.GetList(_web);

            _web.EnsureProperties(wctx => wctx.ServerRelativeUrl);
            _list.EnsureProperties(lctx => lctx.Fields.Include(lctxi => lctxi.InternalName, lctxi => lctxi.JSLink));


            string webRelativeUrl = _web.ServerRelativeUrl;
        
            var fieldToMod = _list.Fields.FirstOrDefault(fod => fod.InternalName == FieldIdentity.Name);
            if (fieldToMod != null)
            {
                if (ShouldProcess(string.Format("Processing field {0} with JSLINK {1}", FieldIdentity.Name, JsLink)))
                {
                    fieldToMod.JSLink = JsLink;
                    fieldToMod.Update();
                    ClientContext.ExecuteQueryRetry();
                }
            }
        }
    }
}