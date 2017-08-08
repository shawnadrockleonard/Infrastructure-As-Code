
## List [Call Out Link] Operations
The following will reverse any stored call out document URL.  
The sample below will query a document library for invalid callout URLs.  
Once the URL is found it will export the objects.  
You can further refine the exported files and pass the callout model collection into the next cmdlet. 
The CSOM will update the docurl and use SystemUpdate to maintain the Modified By and Modified Date.
Once this process is complete your migrated documents should launch in the browser or the appropriate client.


### Callout screenshot
<img src="https://raw.githubusercontent.com/pinch-perfect/Infrastructure-As-Code/master/HowToExtend/imgs/call-out-links.PNG" />

#### Powershell Example


```posh

$siteurl = "https://[tenant].sharepoint.com/sites/site1"
$list = "Sample List"
$migratedurl = "onpremhostheader"
$csvpath = "C:\vData\invalidlinks.csv"

# Root Folder
$path = "/sites/site1/Sample_List"'

# Root/Folder
$path = "/sites/site1/Sample_List/Folder1"'

# Root/Folder/Sub Folder
$path = "/sites/site1/Sample_List/Folder2/Subfolder1"'



# Connect and claim a client context
Connect-SPIaC -Url $siteurl -CredentialName "sponline"



# Find all Call Out Links with a URL that matches the on-premises host name or previous location
$invalidlinks = Find-IaCCallOutLinks -List $list -PartialUrl $migratedurl -Path $path -EndId 100000 -Verbose

# Write the collection of results to a CSV file (optional)
$invalidlinks | Export-Csv -Path $csvpath


# This isn't necessary but you can use this as a pre-parser
$importedcsv = Import-Csv $csvpath
$items = $importedcsv | Where-Object { $_.DocIdUrl -ilike '*' + $migratedurl + '*' }


# Retrieve List and Update View
Set-IaCCallOutLinksByObjects -List $list -Items $items -PartialUrl $migratedurl -Verbose



# Disconnect the client context
Disconnect-SPIaC


```
