using InfrastructureAsCode.Core.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace InfrastructureAsCode.Core.HttpServices
{
    /// <summary>
    /// Control access request settings.
    /// <remarks>Some of the settings are only valid in the Office365</remarks>
    /// </summary>
    public class HttpRequestAccessRequestSettings : HttpRemoteOperation
    {
        #region CONSTRUCTORS

        public HttpRequestAccessRequestSettings(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "")
            : base(TargetUrl, authType, User, Password, Domain)
        {
            this.resetUrl = "";

            //his.EmailAddresses = "x@email.com";
        }

        #endregion

        #region PROPERTIES

        public override string OperationPageUrl
        {
            get
            {
                // return "/_layouts/15/reghost.aspx?type=web&Source=%2F%5Flayouts%2F15%2Fuser%2Easpx&IsDlg=1";
                return this.resetUrl;
            }
        }
        public string resetUrl
        {
            get;
            set;
        }
        public string pageContent
        {
            get;
            set;
        }



        #endregion

        #region METHODS

        public override void AnalyzeRequestResponse(string page)
        {
            pageContent = page;
        }

        public override void SetPostVariables()
        {
            // Set operation specific parameters
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$ctl01$RptControls$BtnReset");

            Console.WriteLine(" - Turning AllowMembersToShareSite flag on -- " + resetUrl);
            this.PostParameters.Add("ctl00_PlaceHolderMain_ctl00_ctl04_TxtUrl", resetUrl);
        }

        #endregion

    }
}
