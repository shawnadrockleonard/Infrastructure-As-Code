
# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Read the JSON file and process Custom Actions
Set-IaCCustomAction -FilePath "c:\filedir\sample.json" -Verbose


# Close the connection
Disconnect-SPIaC