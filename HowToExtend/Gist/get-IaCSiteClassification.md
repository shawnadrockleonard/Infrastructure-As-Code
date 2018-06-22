# Site Classifications

https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification provides a great overview of the site classification feature and how to enable/update/disable in your tenant.
This is a simple discussion on registering an application Azure AD v2 App.
A cmdlet to interogate the site classifications.


At a minimum you'll need the following from your App Registration:
- TenantId 
The specified Azure AD Tenant Unique ID

- MSALClientID
The Registered App Unique ID

- MSALClientSecret
The Registered App Unique ID client secret which has a specified expiration date

- PostLogoutRedirectURI
The Registered App redirect URI.  Typically used in an authentication redirect scenario but if using ConfidentialClient will be used to claim a token



```posh

# Connect to your Site or Tenant URL
    Connect-SPIaC -url "https://[tenant].sharepoint.com" -CredentialName "[credential name]"


# Call the Site Classifications

    Get-IaCSiteClassifications -TenantId $tenantId -MSALClientID $msalID -MSALClientSecret $msalSecret -PostLogoutRedirectURI $msalRedirectUri -Verbose


# Disconnect the session
    Disconnect-SPIaC

```