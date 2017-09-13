$graphparms = @{
appid='<client id>';
appsecret='<client secret>';
aaddomain='<domain>.onmicrosoft.com';
ResourceUri='https://graph.microsoft.com';
}


Report-IaCUsageAnalytics @graphparms -ReportType OneDriveActivity -ViewType Users -Period D30 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType OneDriveActivity -ViewType Files -Period D30 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType OneDriveActivity -ViewType Detail -Date ([System.DateTime]::Parse("2017-09-06")) -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType OneDriveUsage -ViewType Account -Period D30 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType OneDriveUsage -ViewType Files -Period D30 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType OneDriveUsage -ViewType Storage -Period D30 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType OneDriveUsage -ViewType Detail -Date ([System.DateTime]::Parse("2017-09-06")) -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType SharePointActivity -ViewType Users -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType SharePointActivity -ViewType Pages -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType SharePointActivity -ViewType Files -Period D7 -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType SharePointSiteUsage -ViewType Sites -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType SharePointSiteUsage -ViewType Pages -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType SharePointSiteUsage -ViewType Storage -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType SharePointSiteUsage -ViewType Files -Period D7 -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType EmailActivity -ViewType Activity -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType EmailActivity -ViewType Users -Period D7 -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType Office365GroupsActivity -ViewType Activity -Period D7 -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType Office365ActiveUsers -ViewType Services -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType Office365ActiveUsers -ViewType Users -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType Office365ActiveUsers -ViewType Detail -Date ([System.DateTime]::Parse("2017-09-06")) -Verbose

Report-IaCUsageAnalytics @graphparms -ReportType MailboxUsage -ViewType Mailbox -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType MailboxUsage -ViewType Storage -Period D7 -Verbose
Report-IaCUsageAnalytics @graphparms -ReportType MailboxUsage -ViewType Quota -Period D7 -Verbose
