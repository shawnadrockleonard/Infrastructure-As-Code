# Office 365 Usage and Analytics

The following cmdlet and associated .net code helps provide a framework upon which you can build or consume Usage and Analytics.
The cmdlet supports the SwitchParameter BetaEndPoint which will execute the Usage API's with the application/json format.  
You can test the result and speed.  Otherwise it defaults to CSV download which does not support paging and pulls the entire data set.

## Working with Office 365 usage reports in Microsoft Graph
With Microsoft Graph, you can access Office 365 usage reports resources to get the information about how people in your business are using Office 365 services. 
For example, you can identify who is using a service a lot and reaching quotas, or who may not need an Office 365 license at all.
- https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report

### Authorization
Microsoft Graph controls access to resources via permissions. You must specify the permissions you need in order to access Reports resources. 
Typically, you specify permissions in the Azure Active Directory (Azure AD) portal. For more information, see Microsoft Graph permissions reference and Reports permissions.



#### Powershell usage

```posh

# Establish the Azure AD Credentials

$graphparms = @{
	AppId='<client id>';
	AppSecret='<client secret>';
	AADDomain='<domain>.onmicrosoft.com';
	ResourceUri='https://graph.microsoft.com';
}

# Directory to which CSVs will be written
	$dir = "L:\temp\usage\"

# arbitrary date
	$reportDate = (Get-Date).AddDays(-20)

# Connect using the Azure AD Application Details
	Connect-IaCADALv1 @graphparms -Verbose


# use the v1.0 endpoint
	Report-IaCUsageAnalytics -ReportType NONE -ReportUsageType NONE -Date $reportDate -Period D30 -Verbose

# use the beta endpoint
	Report-IaCUsageAnalytics -ReportType NONE -ReportUsageType NONE -Date $reportDate -Period D30 -BetaEndPoint -Verbose


# Office 365 Groups
	Report-IaCUsageAnalytics -ReportType Office365Groups -ReportUsageType  getOffice365GroupsActivityActivity  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose

# Office 365 
	Report-IaCUsageAnalytics -ReportType Office365 -ReportUsageType  getOffice365ActiveUsersServices  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Office365 -ReportUsageType  getOffice365ActiveUsersUsers  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Office365 -ReportUsageType  getOffice365ActiveUsersDetail -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	
# OneDrive 
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveActivityFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveUsageAccountCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveUsageFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveUsageStorage -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType OneDrive -ReportUsageType  getOneDriveUsageAccountDetail -DataDirectory $dir -Date $reportDate -Period D30


# SharePoint 
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointActivityFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointActivityPages -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointSiteUsageDetail -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointSiteUsageFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointSiteUsagePages -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointSiteUsageSiteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType SharePoint -ReportUsageType  getSharePointSiteUsageStorage -DataDirectory $dir -Date $reportDate -Period D30


# Exchange 
	Report-IaCUsageAnalytics -ReportType Exchange -ReportUsageType  getEmailActivityActivity  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Exchange -ReportUsageType  getEmailActivityUsers  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Exchange -ReportUsageType  getMailboxUsageMailbox  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Exchange -ReportUsageType  getMailboxUsageStorage  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType Exchange -ReportUsageType  getMailboxUsageQuota  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose

# Skype for Business

	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessDeviceUsageDistributionUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessDeviceUsageUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessDeviceUsageUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessOrganizerActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessOrganizerActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessOrganizerActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessParticipantActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessParticipantActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessParticipantActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessPeerToPeerActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessPeerToPeerActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType Skype -ReportUsageType  getSkypeForBusinessPeerToPeerActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30


```