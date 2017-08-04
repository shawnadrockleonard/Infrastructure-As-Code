using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.Sites
{
    /// <summary>
    /// The function cmdlet will allow you to specify a JSON file and update the Site or Web with the appropriate User Custom Actions
    /// </summary>
    /// <remarks>
    /// Set-IaCCustomActionByXm -XmlFilePath c:\file\CustomElement.xml
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCCustomActionByXm", SupportsShouldProcess = true)]
    public class SetIaCCustomActionByXml : IaCCmdlet
    {
        /// <summary>
        /// The list to which the custom action will be added
        /// </summary>
        [Parameter(Mandatory = false)]
        public ListPipeBind Identity { get; set; }

        /// <summary>
        /// The full file path to the XML file
        /// </summary>
        [Parameter(Mandatory = true)]
        public string XmlFilePath { get; set; }

        /// <summary>
        /// Validate file path
        /// </summary>
        protected override void OnBeginInitialize()
        {
            var fileInfo = new System.IO.FileInfo(XmlFilePath);
            if (!fileInfo.Exists)
            {
                throw new System.IO.FileNotFoundException("File not found", fileInfo.Name);
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var fileInfo = new System.IO.FileInfo(XmlFilePath);
            var site = this.ClientContext.Site;
            var web = this.ClientContext.Web;
            this.ClientContext.Load(site, ccsu => ccsu.ServerRelativeUrl, cssu => cssu.UserCustomActions);
            this.ClientContext.Load(web, ccwu => ccwu.ServerRelativeUrl, ccwu => ccwu.UserCustomActions);
            this.ClientContext.ExecuteQueryRetry();

            var siteurl = TokenHelper.EnsureTrailingSlash(site.ServerRelativeUrl);
            var weburl = TokenHelper.EnsureTrailingSlash(web.ServerRelativeUrl);


            if (!string.IsNullOrEmpty(Identity.Title))
            {
                var thislist = web.GetListByTitle(Identity.Title);

                XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

                if (!string.IsNullOrEmpty(XmlFilePath))
                {
                    var xdoc = XDocument.Load(XmlFilePath);
                    var customActionNode = xdoc.Element(ns + "Elements").Element(ns + "CustomAction");
                    var customActionName = customActionNode.Attribute("Id").Value;
                    var commandUIExtensionNode = customActionNode.Element(ns + "CommandUIExtension");
                    var xmlContent = commandUIExtensionNode.ToString();
                    var location = customActionNode.Attribute("Location").Value;
                    var sequence = 1000;
                    if (customActionNode.Attribute("Sequence") != null)
                    {
                        sequence = Convert.ToInt32(customActionNode.Attribute("Sequence").Value);
                    }
                    thislist.AddOrUpdateCustomActionLink(customActionName, xmlContent, location, sequence);
                }
            }
        }
    }
}
