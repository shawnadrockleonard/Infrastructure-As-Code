# Set List View Details
Demonstrates how to update details of a List View


```posh

# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Update the JSLink
$listName = "List Name"
$viewName = "View Name"
jsLinks = @("~sitecollection/SiteAssets/JS/file1.js", "~sitecollection/SiteAssets/JS/file2.js")

Set-IaCListViewMinimal -List $listName -Identity $viewName -JsLinkUris $jsLinks -Verbose



# Update the ViewCAML

$listName = "List Name"
$viewName = "View Name"
$viewquery = "<Query></Query>"

Set-IaCListViewMinimal -List $listName -Identity $viewName -CamlQuery $viewquery -Verbose



# Close the connection
Disconnect-SPIaC

```