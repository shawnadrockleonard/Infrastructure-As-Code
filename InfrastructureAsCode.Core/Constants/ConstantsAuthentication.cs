using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Constants
{
    /// <summary>
    /// Used to establish an oAuth v1 or v2 token
    /// </summary>
    public static class ConstantsAuthentication
    {
        /// <summary>
        /// Common end-point for Microsoft Online Services. You should no longer use https://login.windows.net
        /// </summary>
        public const string CommonAuthority = "https://login.microsoftonline.com/common/";

        /// <summary>
        /// Endpoint for the Microsoft Azure AD endpoint
        /// </summary>
        public const string GraphServiceUrl = "https://graph.windows.net";


        public const string O365UnifiedAPIResource = @"https://graph.microsoft.com/";


        internal const string ActiveDirectoryAuthenticationServiceUrl = "https://login.microsoftonline.com/common/oauth2/authorize";

        internal const string ActiveDirectorySignOutUrl = "https://login.microsoftonline.com/common/oauth2/logout";

        internal const string ActiveDirectoryTokenServiceUrl = "https://login.microsoftonline.com/common/oauth2/token";

        public const string NameClaimType = "name";

        public const string IssuerClaim = "iss";

        public const string TenantAuthority = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";

        public const string Authority = "https://login.microsoftonline.com/common/v2.0/";

        public const string RedirectUri = "https://localhost:44321/";

        public const string TenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";

        public const string MicrosoftGraphGroupsApi = "https://graph.microsoft.com/v1.0/groups";

        public const string MicrosoftGraphUsersApi = "https://graph.microsoft.com/v1.0/users";

        public const string AdminConsentFormat = "https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&state={2}&redirect_uri={3}";

        public const string MSGraphScope = "https://graph.microsoft.com/.default";

    }

}
