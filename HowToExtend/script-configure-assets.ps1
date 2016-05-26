# Connect-SPIaC -Url "https://[tenant].sharepoint.com" -UserName "[user]@[tenant].onmicrosoft.com" # open Connection

# Location of my SiteAssets folder which contains my JS, XSL, CSS, and other artifacts
$RelativeOrFullPath = "[REPO Path]\development-samples\Sample02\AppFiles"

Set-IaCProvisionAssets -SiteContent $RelativeOrFullPath -SiteActionFile "js" -Verbose # Uploads JS

Set-IaCProvisionAssets -SiteContent $RelativeOrFullPath -SiteActionFile "html" -Verbose # Uploads HTML

Set-IaCProvisionAssets -SiteContent $RelativeOrFullPath -SiteActionFile "css" # Uploads everything in the CSS folder

Set-IaCProvisionAssets -SiteContent $RelativeOrFullPath -SiteActionFile "xsl" -Verbose # SUploads XSL

Set-IaCProvisionAssets -SiteContent $RelativeOrFullPath # uploads all SiteAssets

Disconnect-SPIaC # close the ClientContext cache