
# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Read the XML file and add/update Custom Actions
Set-IaCCustomActionByXml -Identity "List Title" -XmlFilePath "c:\filedir\sample.xml" -Verbose


# Close the connection
Disconnect-SPIaC