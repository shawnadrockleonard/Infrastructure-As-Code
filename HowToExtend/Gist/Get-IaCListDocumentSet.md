# Retreives List Items from Document Sets

Processes a list and queries Document sets

- List "List pipe band"
- View "View pipe band"
- TargetLocation "The absolute or relative path to a directory where folders/files will be written"


```posh

# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Scan the list and process via Document Sets
$listname = "List 1"
$viewname = "View 1"
$targetdir = "c:\temp\docSets"

Get-IaCListDocumentSet -List $listname -View $viewname -TargetLocation $targetdir -Verbose


# Close the connection
Disconnect-SPIaC

```