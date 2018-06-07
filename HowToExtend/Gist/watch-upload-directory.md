# The Directory Watcher
The following cmdlet demonstrates how you can monitor a directory for file changes.
You specify to whcih site and library the files should be uploaded.  
The watcher will recursively watch/read the files in directories and emulate the directory structure in the target list.

## Parameters

- SiteContent
The folder where the watcher will periodically poll the content changes

- Watch
Tells the cmdlet to continously poll the directory

- Target
The document library where the files will be uploaded

- FileNameFilters
The array of wildcard file extensions to upload.  If none specified it will upload all files. 

- CompareDateTime
Specified if you want to target a file modified date otherwise it will default to DateTime.Now

-WaitSecond
Poll interval





```posh

# Connect to the SharePoint Site
Connect-SPIaC -Url "https://[tenant].sharepoint.com/sites/SiteA" -CredentialName "spiqonline"


# Watch for any JS Changes
$dtcompare = $null
$dtcompare = ([System.DateTime]::Now.AddMinutes(-5))
Watch-IaCDirectoryAndUpload -SiteContent ".\Apps\SiteAssets" -Watch -TargetList "Site Assets" -FileNameFilters @("*.js","*.xsl") -CompareDatetime $dtcompare -WaitSeconds 1 -Verbose
```
