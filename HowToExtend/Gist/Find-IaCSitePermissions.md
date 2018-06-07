# Scans for a user

The following cmdlet will recursively scan a site or web looking for the particular username, group, or identity claim.
This cmdlet requires Global Administrator or SharePoint Administrator authorization


```posh

# Connect to site
$siteurl = "https://[tenant]-admin.sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Scan the URL recursively for the username
Find-IaCSitePermissions -SiteUrl $url -UserName "userB@domain.com" -Verbose


# Close the connection
Disconnect-SPIaC

```