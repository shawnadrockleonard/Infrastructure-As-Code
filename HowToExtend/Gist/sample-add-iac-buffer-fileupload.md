# Markdown File



### Uploading a large file in chunks
This is derived from an original sample at (https://github.com/SharePoint/PnP/tree/dev/Samples/Core.LargeFileUpload#large-file-handling---option-3-startupload-continueupload-and-finishupload)

#### Powershell Example


```posh

$siteurl = "https://<tenant>-my.sharepoint.com/personal/<username>"

# Connect and claim a client context
Connect-SPIaC -Url $siteurl -CredentialName "spadmin"

# Grab a reference to the Document Library
$list = Get-IaCList -Identity "Documents" -Verbose

# 90 Mbs file with full path
$file = C:\<FileDirectoryPath>\LargeFile.csv

# Calls the File upload sample specifying the folder structure
Add-IaCBufferFileUpload -ListTitle $list -FileName $file -FolderName "Cleanup/LargeFiles" -Clobber -Verbose


# Disconnect the context
Disconnect-SPIaC

```

The resulting output should look like this in Powershell.  This is uploading in 8MB chunks

```text

:>Uploading file to https://<tenant>-my.sharepoint.com/personal/<username>/Documents/Cleanup/LargeFiles
:>File length 92MB uploading with buffer context LargeFile.csv
:>File length 92MB uploading with synchronous context LargeFile.csv => Started
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 8MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 16MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 24MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 32MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 40MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 48MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 56MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 64MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 72MB
:>File length 92MB uploading with synchronous context LargeFile.csv fileoffset => 80MB
:>File length 92MB uploading with synchronous context LargeFile.csv FinishUploading => 88
:>Successfully uploaded C:\<FileDirectoryPath>\LargeFile.csv

```