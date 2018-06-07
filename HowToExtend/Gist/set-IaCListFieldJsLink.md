# Update a SPField JSLink

Demonstrates how to configure a JSLink on a Field



```posh

# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


#
$customlist = "List Title"
#
Set-IaCListFieldJsLink -Identity $customlist -FieldIdentity "Field_x0020_Name" -JsLink "~site/SiteAssets/js/csrfile.js"


# Close the connection
Disconnect-SPIaC

```