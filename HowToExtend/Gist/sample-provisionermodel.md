# Markdown File



### Callout screenshot


#### Powershell Example


```posh

$siteurl = "https://[tenant].sharepoint.com/sites/site1"
$list = "Sample List"
$jsonpath = "C:\vData\site-provisioner.json"


# Connect and claim a client context
Connect-SPIaC -Url $siteurl -CredentialName "sponline"

# Call the site to build a site definition file
Get-IaCProvisionResources -ProvisionerFilePath $jsonpath -SpecificListName $list -Verbose


# Disconnect the context
Disconnect-SPIaC

```