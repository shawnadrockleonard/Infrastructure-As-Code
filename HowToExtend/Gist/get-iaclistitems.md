# Demonstrates how to Query a List
This can also demonstrate how to overcome list threshold issues


### Sample
```posh

Connect-SPIaC -Url "https://[tenant].sharepoint.com/sites/siteA" -CredentialName "spiqonline"
$ctx = Get-IaCSPContext

$list = $ctx.Web.Lists.GetByTitle("ListName")
$ctx.Load($list)
$ctx.ExecuteQuery()


Get-IaCListItems `
    -Query '<View Scope="Recursive"><Query><Where><And><And><And><And><Eq><FieldRef Name="SiteCollectionName" LookupId="TRUE"/><Value Type="Lookup">7</Value></Eq><Eq><FieldRef Name="SiteName" /><Value Type="Text">PIOneers</Value></Eq></And><Eq><FieldRef Name="TypeOfSiteID"/><Value Type="Text">0</Value></Eq></And><Neq><FieldRef Name="ID"/><Value Type="Integer">0</Value></Neq></And><Neq><FieldRef Name="RequestRejectedFlag"/><Value Type="Text">Yes</Value></Neq></And></Where></Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="SiteURL" /></ViewFields></View>' `
    -List $list `
    -PageSize 5


$list = $ctx.Web.Lists.GetByTitle("ListName")
$ctx.Load($list)
$ctx.ExecuteQuery()


Get-IaCListItems `
    -Query '<View Scope="Recursive"><Query><Where><And><Eq><FieldRef Name="CollectionSiteType" LookupId="TRUE"/><Value Type="Lookup">7</Value></Eq><Neq><FieldRef Name="RequestCompletedFlag"/><Value Type="Text">No</Value></Neq></And></Where></Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Title1" /></ViewFields></View>' `
    -List $list `
    -PageSize 5

# Test In Clause
Get-IaCListItems `
    -Query '<View Scope="Recursive"><Query><Where><And><Or><Eq><FieldRef Name="CollectionSiteType" LookupId="TRUE"/><Value Type="Lookup">7</Value></Eq><Eq><FieldRef Name="CollectionSiteType" LookupId="TRUE"/><Value Type="Lookup">9</Value></Eq></Or></Where></Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Title1" /></ViewFields></View>' `
    -List $list `
    -PageSize 5


```