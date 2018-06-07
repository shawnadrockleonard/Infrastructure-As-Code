# Sample to sync a directory to a SharePoint Library
Provides sample code AS-IS to sync a directory and all of its subfolders and files into a document library.
Will recursively enumerate the directory and build the folder structure in SharePoint.


### Parameters
- Library "List name, GUID, or List object"
- DirectoryPath "The full or relative path to a directory that will sync"
- MetaDataCsvFile "OPTIONAL: CSV that contains filename and "tags" which will upload as metadata"


```posh

# Connect to site
$siteurl = "https://[tenant].sharepoint.com/sites/[siteurl]"

Connect-SPIaC -CredentialName "sponline" -Url $siteurl


# Syncs a directory to a SharePoint list
$listname = "List 1"
$syncdir = "c:\temp\docSets"

Set-IaCSyncDirectory -Library $listname -DirectoryPath $syncdir -Verbose


# Syncs a directory to a SharePoint list
$listname = "List 1"
$syncdir = "c:\temp\docSets"
$csv = "c:\temp\metadata.csv"

Set-IaCSyncDirectory -Library $listname -DirectoryPath $syncdir -MetaDataCsvFile $csv -Verbose


# Close the connection
Disconnect-SPIaC

```