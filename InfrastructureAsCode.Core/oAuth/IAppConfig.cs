using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// Represents the Application Config AppSettings
    /// </summary>
    public interface IAppConfig
    {
        /// <summary>
        /// Gets or sets the PostLogoutRedirectURI for Active Directory authentication. The Post Logout Redirect Uri is the URL where the user will be redirected after they have signed out
        /// </summary>
        string PostLogoutRedirectURI { get; }

        /// <summary>
        /// Gets or sets the application ID for Active Directory authentication. The Client ID is used by the application to uniquely identify itself to Azure AD.
        /// </summary>
        string ClientID { get; }

        /// <summary>
        /// Gets or sets the client secret for Active Directory authentication. The ClientSecret is a credential used to authenticate the application to Azure AD.  Azure AD supports password and certificate credentials.
        /// </summary>
        string ClientSecret { get; }

        /// <summary>
        /// Gets or sets the Tenant Domain
        /// </summary>
        string TenantDomain { get; }

        /// <summary>
        /// Gets or sets the Tenant Id
        /// </summary>
        string TenantId { get; }

        /// <summary>
        /// Gets or sets if the Application is Multi-Tenant
        /// </summary>
        bool? IsAppMultiTenent { get; }

        /// <summary>
        /// TODO
        /// </summary>
        string ServiceResource { get; }


        string Audience { get; }


        string SPClientID { get; }

        string SPClientSecret { get; }


        string MSALClientID { get; }

        string MSALClientSecret { get; }

        string MSALScopes { get; }
    }
}
