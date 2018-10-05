using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class OD4BSiteCollectionModel : CollectionModel
    {
        public OD4BSiteCollectionModel()
        {
        }

        public string PersonalSpaceProperty { get; set; }

        public string NameProperty { get; set; }

        public string UserName { get; set; }

        public string PictureUrl { get; set; }

        public string AboutMe { get; set; }

        public string SpsSkills { get; set; }

        public string Manager { get; set; }

        public string MailingZipCode { get; set; }

        public string WorkPhone { get; set; }

        public string Department { get; set; }

        public string Company { get; set; }

        public string OWAUrl { get; internal set; }

        public string ProxyAddresses { get; internal set; }

        public string MySiteUpgrade { get; internal set; }

        public string PrivacyActivity { get; internal set; }

        public string PrivacyPeople { get; internal set; }

        public string EmailOptin { get; internal set; }

        public string Locale { get; internal set; }

        public string TimeZone { get; internal set; }

        public string AccountName { get; internal set; }

        public string FirstName { get; internal set; }

        public string HireDate { get; internal set; }

        public string Assistant { get; internal set; }

        public string JobTitle { get; internal set; }

        public string Education { get; internal set; }

        public string WebSite { get; internal set; }

        public string School { get; internal set; }

        public string DistinguishedName { get; internal set; }

        public string LastName { get; internal set; }

        public string UserPrincipalName { get; internal set; }

        public string Title { get; internal set; }

        public string WorkEmail { get; internal set; }

        public string HomePhone { get; internal set; }

        public string CellPhone { get; internal set; }

        public string Office { get; internal set; }

        public string Location { get; internal set; }

        public string Fax { get; internal set; }

        public string MailingAddress { get; internal set; }
    }
}
