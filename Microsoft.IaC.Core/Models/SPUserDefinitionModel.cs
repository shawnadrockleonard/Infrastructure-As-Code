using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IaC.Core.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace IaC.Core.Models
{
    /// <summary>
    /// Represents a user object or user profile
    /// </summary>
    public class SPUserDefinitionModel
    {
        public Guid Id { get; set; }

        public string UserName { get; set; }

        public string UserEmail { get; set; }

        public string UserDisplay { get; set; }

        public string Organization { get; set; }

        public string Manager { get; set; }

        public string OrganizationAcronym { get; set; }

        public string OD4BUrl { get; set; }

        public int? UserIndex { get; set; }

        /// <summary>
        /// CSOM Info
        /// </summary>
        public UserIdInfo UserId { get; set; }

        /// <summary>
        /// CSOM Principal Type
        /// </summary>
        public PrincipalType PrincipalType { get; set; }

        /// <summary>
        /// Object ID from Azure AD
        /// </summary>
        public int GuidId { get; set; }

        /// <summary>
        /// The latest post in the profile
        /// </summary>
        public string LatestPost { get; set; }

        /// <summary>
        /// User Profile object identifier
        /// </summary>
        public string UserProfileGUID { get; set; }
        public string SPSDistinguishedName { get; set; }
        public string SPSSid { get; set; }
        public string MSOnlineObjectId { get; set; }
    }
}
