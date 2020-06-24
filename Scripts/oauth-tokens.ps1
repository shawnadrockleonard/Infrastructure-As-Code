<#
    .Synopsis
        Get access token for AAD web app.

    .Description
        Authorizes AAD app and retrieves access token using OAuth 2.0 and endpoints.
        Refreshes the token if within 5 minutes of expiration or, optionally forces refresh.
        Sets global variable ($Global:accessTokenResult) that can be used after the script runs.

    .Todo
        Add ability to handle refresh token input and access token retrieval without re-authorization.

    .Example 
        The following returns the access token result from AAD with admin consent authorization and caches the result.

        PS> .\aad_web.ps1 -Clientid "" -Clientsecret "" -Resource "https://TENANT.sharepoint.com" -Redirecturi "https://localhost:44385" -Scope "" -AdminConsent -Cache
    
    .Example 
        The following returns the access token result from AAD with admin consent authorization or refreshes the token.

        PS> .\aad_web.ps1 -Clientid "" -Clientsecret "" -Resource "https://TENANT.sharepoint.com" -Redirecturi "https://localhost:44385" -Scope "" -AdminConsent
    
    .Example 
        The following returns the access token result from AAD or from cache, forces refresh so the token is good for an hour and outputs to a file

        PS> .\aad_web.ps1 -Clientid "" -Clientsecret "" -Resource "https://TENANT.sharepoint.com" -Redirecturi "https://localhost:44385" -Scope "" -Refresh Force | Out-File c:\temp\token.txt

    .PARAMETER ClientId 
        The AAD App client id.
    .PARAMETER ClientSecret
        The AAD App client secret.	
    .PARAMETER RedirectUri
        The redirect uri configured for that app.
    .PARAMETER Resource
        The resource the app is attempting to access (i.e. https://TENANT.sharepoint.com)
    .PARAMETER Scope
        Permission scopes for the app (optional).
    .PARAMETER AdminConsent
        Will perform admin consent (optional).
    .PARAMETER Cache
        Cache the access token in the temp directory for subsequent retrieval (optional).
    .PARAMETER Refresh
        Options (Yes, No, Force). Will automatically enabling caching if "Yes" or "Force" are used.
        Yes: Refresh token if within 5 minutes of expiration if cached token found.
        No: Do not refresh and re-authorize.
        Force: Forfce a refresh if cached token found.

#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $true)]
    [string]$RedirectUri,
    [Parameter(Mandatory = $true)]
    [string]$Resource,
    [Parameter(Mandatory = $false)]
    [string]$Scope,
    [Parameter(Mandatory = $false)]
    [switch]$AdminConsent,
    [Parameter(Mandatory = $false)]
    [switch]$Cache,
    [Parameter(Mandatory = $false)]
    [ValidateSet("Yes", "No", "Force")]
    [ValidateNotNullOrEmpty()]
    [string]$Refresh = "Yes"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

$isCache = $Cache.IsPresent
$isRefresh = (($Refresh -eq "Yes") -or ($Refresh -eq "Force"))
$refreshForce = $Refresh -eq "Force"

if ($isRefresh) {
    $isCache = $true
}

# Don't edit variables below (unless there's a bug)
$clientSecretEncoded = [uri]::EscapeDataString($clientSecret)
$redirectUriEncoded = [uri]::EscapeDataString($redirectUri)
$resourceEncoded = [uri]::EscapeDataString($resource)
$accessTokenUrl = "https://login.microsoftonline.com/common/oauth2/token"
$cacheFilePath = [System.IO.Path]::Combine($env:TEMP, "aad_web_cache_$clientId.json")

$accessTokenResult = $null
$adminConsentText = ""
if ($adminConsent) {
    $adminConsentText = "&prompt=admin_consent"
}

$authorizationUrl = "https://login.microsoftonline.com/common/oauth2/authorize?resource=$resourceEncoded&client_id=$clientId&scope=$scope&redirect_uri=$redirectUriEncoded&response_type=code$adminConsentText"

function Invoke-OAuth() {
    $Global:authorizationCode = $null

    $form = New-Object Windows.Forms.Form
    $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
    $form.Width = 640
    $form.Height = 480
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $web = New-Object Windows.Forms.WebBrowser
    $form.Controls.Add($web)

    $web.Size = $form.ClientSize
    $web.DocumentText = "<html><body style='text-align:center;overflow:hidden;background-image:url(https://secure.aadcdn.microsoftonline-p.com/ests/2.1.6856.20/content/images/backgrounds/0.jpg?x=f5a9a9531b8f4bcc86eabb19472d15d5)'><h3 id='title'>Continue with current user or logout?</h3><div><input id='cancel' type='button' value='Continue' /></div><br /><div><input id='logout' type='button' value='Logout' /></div><h5 id='loading' style='display:none'>Working on it...</h5><script type='text/javascript'>var logout = document.getElementById('logout');var cancel = document.getElementById('cancel');function click(element){document.getElementById('title').style.display='none';document.getElementById('loading').style.display='block';logout.style.display='none';cancel.style.display='none';if (this.id === 'logout'){window.location = 'https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=' + encodeURIComponent('$authorizationUrl');}else{window.location = '$authorizationUrl';}}logout.onclick = click;cancel.onclick = click;</script></body></html>"

    $web.add_DocumentCompleted(
        {
            $uri = [uri]$redirectUri
            $queryString = [System.Web.HttpUtility]::ParseQueryString($_.url.Query)

            if ($_.url.authority -eq $uri.authority) {
                $authorizationCode = $queryString["code"]
        
                if (![string]::IsNullOrEmpty($authorizationCode)) {
                    $form.DialogResult = "OK"
                    $Global:authorizationCode = $authorizationCode
                    $Global:authorizationCodeTime = [datetime]::Now
                }

                $form.close()
            }
        })

    $dialogResult = $form.ShowDialog()

    if ($dialogResult -eq "OK") {
        $authorizationCode = $Global:authorizationCode
        $headers = @{"Accept" = "application/json;odata=verbose" }
        $body = "client_id=$clientId&client_secret=$clientSecretEncoded&redirect_uri=$redirectUriEncoded&grant_type=authorization_code&code=$authorizationCode"
    
        $accessTokenResult = Invoke-RestMethod -Uri $accessTokenUrl -Method POST -Body $body -Headers $headers
        $Global:accessTokenResult = $accessTokenResult
        $Global:accessTokenResultTime = [datetime]::Now
        $accessTokenResultText = (ConvertTo-Json $accessTokenResult)

        if ($isCache -and ![string]::IsNullOrEmpty($accessTokenResultText)) {
            [void](Set-Content -Path $cacheFilePath -Value $accessTokenResultText)
        }

        Write-Output (ConvertTo-Json $accessTokenResultText)
    }

    $web.Dispose()
    $form.Dispose()
}

function Get-CachedAccessTokenResult() {
    if ($isCache -and [System.IO.File]::Exists($cacheFilePath)) {
        $accessTokenResultText = Get-Content -Raw $cacheFilePath
        if (![string]::IsNullOrEmpty($accessTokenResultText)) {
            $accessTokenResult = (ConvertFrom-Json $accessTokenResultText)
            if (![string]::IsNullOrEmpty($accessTokenResult.access_token)) {
                $Global:accessTokenResult = $accessTokenResult

                return $accessTokenResult
            }
        }
    }

    return $null
}

function Invoke-Refresh() {
    $refreshToken = $accessTokenResult.refresh_token
    $headers = @{"Accept" = "application/json;odata=verbose" }
    $body = "client_id=$clientId&client_secret=$clientSecretEncoded&resource=$resourceEncoded&grant_type=refresh_token&refresh_token=$refreshToken"
    $accessTokenResult2 = Invoke-RestMethod -Uri $accessTokenUrl -Method POST -Body $body -Headers $headers

    $accessTokenResult.scope = $accessTokenResult2.scope
    $accessTokenResult.expires_in = $accessTokenResult2.expires_in
    $accessTokenResult.ext_expires_in = $accessTokenResult2.ext_expires_in
    $accessTokenResult.expires_on = $accessTokenResult2.expires_on
    $accessTokenResult.not_before = $accessTokenResult2.not_before
    $accessTokenResult.resource = $accessTokenResult2.resource
    $accessTokenResult.access_token = $accessTokenResult2.access_token
    $accessTokenResult.refresh_token = $accessTokenResult2.refresh_token

    $Global:accessTokenResult = $accessTokenResult
    $Global:accessTokenResultTime = [datetime]::Now
    $accessTokenResultText = (ConvertTo-Json $accessTokenResult)

    if (![string]::IsNullOrEmpty($accessTokenResultText)) {
        [void](Set-Content -Path $cacheFilePath -Value $accessTokenResultText)
    }

    Write-Output (ConvertTo-Json $accessTokenResultText)
}

$accessTokenResult = Get-CachedAccessTokenResult
if ($accessTokenResult -eq $null) {
    Invoke-OAuth
}
elseif ($refreshForce -or (([datetime]::Parse("1/1/1970")).AddSeconds([int]$accessTokenResult.expires_on).ToLocalTime() -lt ([datetime]::Now).AddMinutes(5))) {
    if ($isRefresh) {
        Invoke-Refresh
    }
    else {
        Invoke-OAuth
    }
}
else {
    Write-Output (ConvertTo-Json $Global:accessTokenResult)
}