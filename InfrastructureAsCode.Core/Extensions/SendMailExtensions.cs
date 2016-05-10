using OfficeDevPnP.Core.Utilities;
using System.Collections.Generic;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Sends an email using the Office 365 SMTP Service
    /// </summary>
    public class SendMailExtensions
    {
        /// <summary>
        /// SMTP Relay EndPoint
        /// </summary>
        public static string Server = "smtp.office365.com";

        public static void ExecuteCmdlet(string servername, string fromAddress, string fromUserPassword, IEnumerable<string> to, IEnumerable<string> cc, string subject, string body, bool sendAsync = false, object asyncUserToken = null)
        {
            MailUtility.SendEmail(Server, fromAddress, fromUserPassword, to, cc, subject, body, sendAsync, asyncUserToken);
        }
    }

}
