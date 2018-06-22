using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Constants
{
    /// <summary>
    /// Used to identify the key which is either in the .config file or Azure config
    /// </summary>
    public static class ConstantsConfigKeys
    {
        public const string CLIENT_ID_KEY = "ida:ClientId";

        public const string CLIENT_SECRET_KEY = "ida:ClientSecret";

        public const string POST_LOGOUTREDIRECTURI_KEY = "ida:PostLogoutRedirectUri";

        public const string TENANT_KEY = "ida:Tenant";

        public const string TENANT_ID_KEY = "ida:TenantId";

        public const string CONNECTOR_URL_KEY = "ConnectorUrl";

        public const string PORTAL_URL_KEY = "PortalUrl";

        public const string NOTIFICATION_INTERVAL_KEY = "NotificationInterval";

        public const string SP_CLIENT_ID_KEY = "ClientId";

        public const string SP_CLIENT_SECRET_KEY = "ClientSecret";

        public const string AUDIENCE_KEY = "ida:Audience";

        public const string MSAL_CLIENT_ID_KEY = "msal:ClientId";

        public const string MSAL_CLIENT_SECRET_KEY = "msal:ClientSecret";

        public const string GRAPH_SCOPES_KEY = "msal:GraphScopes";

    }
}
