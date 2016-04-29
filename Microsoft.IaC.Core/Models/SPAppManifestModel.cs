using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Core.Models
{
    /// <summary>
    /// Represents an App Package manifest
    /// </summary>
    public class SPAppManifestModel
    {
        public SPAppManifestModel()
        {
            this.AppPermissions = new List<SPAppScopePermissionModel>();
        }

        public string Title { get; set; }

        public string AppRequestStatus { get; set; }

        public string AssetId { get; set; }

        public string AppRequestIsSiteLicense { get; set; }

        public IList<SPAppScopePermissionModel> AppPermissions { get; set; }
    }
}
