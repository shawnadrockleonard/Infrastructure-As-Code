# Usage and Adoption reporting

We'll get to configuring and using the Office 365 Usage Reports based on the beta endpoint of the Microsoft Graph API.   First let's quickly overview why this is important.  On October 1st, 2017 some of the API endpoints for the tried and true Reporting webservice https://msdn.microsoft.com/en-us/library/office/jj984325.aspx will be deprecated.   I've grown to depend on this for data modeling and understanding various tenants usage.     Our, Microsoft, march towards a unified API has tremendous benefits for the developer and in reality all platforms.  Azure AD has truly enabled Microsoft to offer first class development and usability.   Enough about that for a moment.  

### o365rwsclient
This is a deprecated class library forked from https://github.com/Microsoft/o365rwsclient 
It has a number of improvements but no effort should be continued forward with this namespace.   

Per public documentation the following is stated:
- With this announcement, we’re starting the deprecation of the following APIs available within the Office 365 Reporting Web Service: ConnectionbyClientType, ConnectionbyClientTypeDetail, CsActiveUser, CsAVConferenceTime, CsP2PAVTime, CsConference, CsP2PSession, GroupActivity, MailboxActivity, GroupActivity, MailboxUsage, MailboxUsageDetail, StaleMailbox and StaleMailboxDetail. We will remove these APIs, as well as any related PowerShell cmdlets, on October 1, 2017.

### Preview API's
All reporting is now available through the Graph API in preview functionality
- https://blogs.office.com/en-us/2017/03/31/whats-new-in-office-365-administration-public-preview-of-microsoft-graph-reporting-apis/
- https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report



So what's replacing the Reporting Service endpoints.  Surely if the reporting service is being deprecated it has a replacement.  That's a bigger conversation but the good news is; you can start building today  https://blogs.office.com/en-us/2017/03/31/whats-new-in-office-365-administration-public-preview-of-microsoft-graph-reporting-apis/     I don't want to take screenshots of how to register an application in Azure AD.   For the sake of this demo I'll send you to 2 locations.

1. Azure AD v1.0 Endpoint - Integrating Applications with Azure AD(https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-integrating-applications)
2. Azure AD v2.0 Endpoint - Register your app(https://developer.microsoft.com/en-us/graph/docs/concepts/auth_register_app_v2)

For these reports you only need v1.0 "Microsoft Graph > Application Permissions > Read all usage reports"  or v2.0 "Application Permissions > Reports.Read.All"    For the sample code I'm providing I'll demonstrate the v1.0 registration and application permissions.  

