
# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


#
$customlist = "List Title"
#
Set-IaCFieldJsLink -Identity $customlist -FieldIdentity "Field_x0020_Name" -JsLink "~site/SiteAssets/js/csrfile.js"


# Close the connection
Disconnect-SPIaC