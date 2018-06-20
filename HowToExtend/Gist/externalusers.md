# External Users

https://support.office.com/en-us/article/external-sharing-overview-c8a462eb-0723-4b0b-8d0a-70feafe4be85

External sharing enables organizations to collaborate.  During the invitation process (if guest links are disabled) an invitation is sent to the recepient.  
The recepient will click the link from their email and accept the invitation.
Typically this action is performed with the identity in which the invitation was received.  
However, it is possible that an end user will create an MSA (Microsoft Account) or accept with an account that is different from the email address to which the invigation was sent.

```
EX:  userB@domain.com accepts the invitation as userB@subdomain.com

"i:0#.f|membership|userB_subdomain.com#ext#@[tenant].onmicrosoft.com" email="userB@domain.com"

```

This get a bit tricky in this situation and the user will potentially find themselves in a denied access or a loop of Access Requests.   
To that end I've added a cmdlet to remove External Users.  There are a few variations here:
1. requires SharePoint Administrator authorization
2. requires Site Collection Administrator authorization



```posh

# We will connect with Stored Credentials to the Tenant Site

Connect-SPIaC -CredentialName "spiqonline" -Url "https://[tenant]-admin.sharepoint.com"

$siteurl = "https://[tenant].sharepoint.com/sites/SiteA"
$username = "userB@domain.com"

# Will query the specific Site and return external users registered with site
Get-IaCExternalUserFromSite -UserName $username -SiteUrl $siteurl -Verbose

# Will query the tenant looking for the specifie username
Get-IaCExternalUserFromSite -UserName $username -Verbose


Remove-SPOUserFromSite -UserName $username -Verbose -WhatIf

Remove-SPOExternalUser -SiteUrl $siteurl -UserName $username -Verbose

# Removes an external user WhatIf demonstrates but does not remove the identity
Remove-SPOExternalUser -SiteUrl $siteurl -UserName $username -Verbose -WhatIf


# Disconnect the context
Disconnect-SPIaC

```