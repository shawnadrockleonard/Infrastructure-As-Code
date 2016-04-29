using IaC.Core.Models;
using IaC.Powershell.CmdLets;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will upgrade the EzForms site specified in the connection to the latest configuration changes
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCProvisionResources")]
    [CmdletHelp("Get site definition components and write to JSON file.", Category = "ETL")]
    public class GetIaCProvisionResources : IaCCmdlet
    {
        [Parameter(Mandatory = false)]
        public string SiteContentPath { get; set; }

        [Parameter(Mandatory = false)]
        public string FileContents { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            if (!System.IO.Directory.Exists(this.SiteContentPath))
            {
                throw new Exception(string.Format("The directory does not exists {0}", this.SiteContentPath));
            }
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            //#TODO: Get configuration from SP Site and write to disk
            var SiteComponents = JsonConvert.DeserializeObject<SiteProvisionerModel>(FileContents);


            //Move away from method configuration into a JSON file
            var filePath = string.Format("{0}\\Content\\{1}", this.SiteContentPath, "Provisioner.json");
            var json = JsonConvert.SerializeObject(SiteComponents, Formatting.Indented);
            System.IO.File.WriteAllText(filePath, json);

            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }
        }
    }
}
