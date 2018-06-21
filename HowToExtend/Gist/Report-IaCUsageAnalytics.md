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
	AppDomain='<domain>.onmicrosoft.com';
	ResourceUri='https://graph.microsoft.com';
}

# Directory to which CSVs will be written
	$dir = "L:\temp\usage\"

# arbitrary date
	$reportDate = (Get-Date).AddDays(-20)


Connect-SPIaC @graphparms -Url "https://[tenant]-admin.sharepoint.com" -SkipTenantAdminCheck


# Office 365 Groups
	Report-IaCUsageAnalytics -ReportType getOffice365GroupsActivityActivity  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose

# Office 365 
	Report-IaCUsageAnalytics -ReportType getOffice365ActiveUsersServices  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getOffice365ActiveUsersUsers  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getOffice365ActiveUsersDetail -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	
# OneDrive 
	Report-IaCUsageAnalytics -ReportType getOneDriveActivityFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getOneDriveActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getOneDriveActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getOneDriveUsageAccountCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getOneDriveUsageFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getOneDriveUsageStorage -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getOneDriveUsageAccountDetail -DataDirectory $dir -Date $reportDate -Period D30


# SharePoint 
	Report-IaCUsageAnalytics -ReportType getSharePointActivityFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointActivityPages -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getSharePointSiteUsageDetail -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointSiteUsageFileCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointSiteUsagePages -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointSiteUsageSiteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSharePointSiteUsageStorage -DataDirectory $dir -Date $reportDate -Period D30


# Exchange 
	Report-IaCUsageAnalytics -ReportType getEmailActivityActivity  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getEmailActivityUsers  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getMailboxUsageMailbox  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getMailboxUsageStorage  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose
	Report-IaCUsageAnalytics -ReportType getMailboxUsageQuota  -DataDirectory $dir -Date $reportDate -Period D30 -Verbose

# Skype for Business

	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessActivityUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessDeviceUsageDistributionUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessDeviceUsageUserCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessDeviceUsageUserDetail -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessOrganizerActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessOrganizerActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessOrganizerActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessParticipantActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessParticipantActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessParticipantActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30

	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessPeerToPeerActivityCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessPeerToPeerActivityMinuteCounts -DataDirectory $dir -Date $reportDate -Period D30
	Report-IaCUsageAnalytics -ReportType getSkypeForBusinessPeerToPeerActivityUserCounts -DataDirectory $dir -Date $reportDate -Period D30


```