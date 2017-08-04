

# Disable the Modern UI for the Site and Subsites
$urls = @("https://[tenant].sharepoint.com/sites/[siteur]/"
)
$urls | ForEach-Object {

# Connect
    $siteurl = $_
    Write-Host ("Now connecting and setting modern UI for {0}" -f $siteurl)
    Connect-SPIaC -CredentialName "sponline" -Url $siteurl

# Set Classic Mode
    Set-IaCCustomAction -Verbose

# Disconnect
    Write-Host ("Now disconnecting and setting modern UI for {0}" -f $siteurl)
    Disconnect-SPIaC
}


# Disable the Modern UI for the Subsite
$weburls = @("https://[tenant].sharepoint.com/sites/[siteurl]/[subsite]/"
)
$weburls | ForEach-Object {

# Connect
    $siteurl = $_
    Write-Host ("Now connecting and setting modern UI for {0}" -f $siteurl)
    Connect-SPIaC -CredentialName "sponline" -Url $siteurl

# Set Classic Mode
    Set-IaCCustomAction -IsWeb -Verbose

# Disconnect
    Write-Host ("Now disconnecting and setting modern UI for {0}" -f $siteurl)
    Disconnect-SPIaC
}