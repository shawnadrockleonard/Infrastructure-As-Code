using InfrastructureAsCode.Core.Constants;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// Handles reading from the configuration file or Azure configuration
    /// </summary>
    internal class AppConfigManager
    {
        /// <summary>
        /// Returns a Domain Model which represents the application config
        /// </summary>
        /// <returns><see cref="AppConfig"/></returns>
        internal AppConfig GetApplicationConfig()
        {
            var _config = new AppConfig
            {
                ClientID = this.ReadConfiguration(ConstantsConfigKeys.CLIENT_ID_KEY),
                ClientSecret = this.ReadConfiguration(ConstantsConfigKeys.CLIENT_SECRET_KEY),
                PostLogoutRedirectURI = this.ReadConfiguration(ConstantsConfigKeys.POST_LOGOUTREDIRECTURI_KEY),
                TenantId = this.ReadConfiguration(ConstantsConfigKeys.TENANT_ID_KEY),
                TenantDomain = this.ReadConfiguration(ConstantsConfigKeys.TENANT_KEY),
                Audience = this.ReadConfiguration(ConstantsConfigKeys.AUDIENCE_KEY),

                SPClientID = this.ReadConfiguration(ConstantsConfigKeys.SP_CLIENT_ID_KEY),
                SPClientSecret = this.ReadConfiguration(ConstantsConfigKeys.SP_CLIENT_SECRET_KEY),

                MSALScopes = this.ReadConfiguration(ConstantsConfigKeys.GRAPH_SCOPES_KEY),
                MSALClientID = this.ReadConfiguration(ConstantsConfigKeys.MSAL_CLIENT_ID_KEY),
                MSALClientSecret = this.ReadConfiguration(ConstantsConfigKeys.MSAL_CLIENT_SECRET_KEY)
            };
            return _config;
        }

        /// <summary>
        /// Stubbed to pull from Azure Storage | Database | File
        /// </summary>
        /// <param name="CONNECTOR_URL_KEY"></param>
        /// <returns></returns>
        private string ReadConfiguration(object CONNECTOR_URL_KEY)
        {
            throw new NotImplementedException();
        }

        #region Private Members

        /// <summary>
        /// Gets the configuration item by a specific key. 
        /// returns <see cref="string.Empty"/> if the value is not set.
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private string ReadConfiguration(string key)
        {
            var _result = string.Empty;

            if (!string.IsNullOrEmpty(key))
            {
                _result = ConfigurationManager.AppSettings[key];
            }
            return _result;

        }

        #endregion
    }
}