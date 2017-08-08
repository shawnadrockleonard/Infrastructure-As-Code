
## List View Operations
The following will render examples on type of CAML Queries and how to update a View in SharePoint.  
Not every query makes sense but it's the progression of a user's requirements over time which can't be implemented through a UI


#### Powershell Example


```posh

# Connect and claim a client context
Connect-SPIaC -Url "https://[tenant].sharepoint.com/sites/site1" -CredentialName "sponline"

$viewxml = "<OrderBy><FieldRef Name='ID' /></OrderBy>
<Where>
<And>
    <And>
        <Or>
            <Eq><FieldRef Name=""Choice_x0020_Field"" /><Value Type=""Text"">TBD</Value></Eq>
            <Eq><FieldRef Name=""Choice_x0020_Field"" /><Value Type=""Text"">value-which-requires-a-user</Value></Eq>
        </Or>
        <Or>
            <Eq><FieldRef Name=""WFStep"" /><Value Type=""Text"">4</Value></Eq>
            <Eq><FieldRef Name=""WFStep"" /><Value Type=""Text"">5</Value></Eq>
        </Or>
    </And>
    <IsNotNull><FieldRef Name=""User_x0020_Field"" /></IsNotNull>
</And>
</Where>"

# Retrieve List and Update View
$list = Get-IaCList -Identity "Sample List" -Verbose
Set-IaCListViewMinimal -List $list -Identity "Sample View" -RowLimit 50 -Verbose

# Get and Update in same call
Set-IaCListViewMinimal -List "Sample List" -Identity "Sample View" -QueryXml $viewxml -Verbose

# Update the JS Link property
Set-IaCListViewMinimal -List "Sample List" -Identity "ViewTitle" -JsLinkUris @("~sitecollection/siteassets/js/jquery.min.js")

# Disconnect the client context
Disconnect-SPIaC

```


###### Field comparisons
This CAML sample performs the following:
- OR
  - WFStep is Step 5
  - AND
    - WFStep is Step 4
    - OR
        1. If a choice field has a specific value then check if a User field is not null; 
        2. If a choice field has a specific TBD value


```xml
<OrderBy><FieldRef Name="ID" /></OrderBy>
<Where>
    <Or>
        <And>
            <Or>
                <And>
                    <IsNotNull><FieldRef Name="User_x0020_Field" /></IsNotNull>
                    <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">value-which-requires-a-user</Value></Eq>
                </And>
                <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">TBD</Value></Eq>
            </Or>
            <Eq><FieldRef Name="WFStep" /><Value Type="Text">4</Value></Eq>
        </And>
        <Eq><FieldRef Name="WFStep" /><Value Type="Text">5</Value></Eq>
    </Or>
</Where>
```


This CAML sample performs the following:
- AND
  - User field is not null
  - OR
    - WFStep is Step 5
    - AND
      - WFStep is Step 4
      - OR
        - Choice Field <> TBD
        - Choice Field = value-which-requires-a-user
```xml
<OrderBy><FieldRef Name="ID" /></OrderBy>
<Where>
<And>
    <Or>
        <And>
            <Or>
            <Neq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">TBD</Value></Neq>
            <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">value-which-requires-a-user</Value></Eq>
            </Or>
            <Eq><FieldRef Name="WFStep" /><Value Type="Text">4</Value></Eq>
        </And>
        <Eq><FieldRef Name="WFStep" /><Value Type="Text">5</Value></Eq>
    </Or>
    <IsNotNull><FieldRef Name="User_x0020_Field" /></IsNotNull>
</And>
</Where>
```


This CAML sample performs the following:
- OR
  - WFStep is Step 5
  - AND
    - WFStep is Step 4
    - OR
      - Choice Field = value-which-requires-a-user
      - Choice Field = TBD
```xml
<OrderBy><FieldRef Name="ID" /></OrderBy>
<Where>
    <Or>
        <And>
            <Or>
                <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">TBD</Value></Eq>
                <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">value-which-requires-a-user</Value></Eq>
            </Or>
            <Eq><FieldRef Name="WFStep" /><Value Type="Text">4</Value></Eq>
        </And>
        <Eq><FieldRef Name="WFStep" /><Value Type="Text">5</Value></Eq>
    </Or>
</Where>
```


This CAML sample performs the following:
- AND
  - User field is not null
  - AND
    - OR
      - WFStep = 4
      - WFStep = 5
    - OR
      - Choice Field = value-which-requires-a-user
      - Choice Field = TBD
```xml
<OrderBy><FieldRef Name="ID" /></OrderBy>
<Where>
    <And>
        <And>
            <Or>
                <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">TBD</Value></Eq>
                <Eq><FieldRef Name="Choice_x0020_Field" /><Value Type="Text">value-which-requires-a-user</Value></Eq>
            </Or>
            <Or>
                <Eq><FieldRef Name="WFStep" /><Value Type="Text">4</Value></Eq>
                <Eq><FieldRef Name="WFStep" /><Value Type="Text">5</Value></Eq>
            </Or>
        </And>
        <IsNotNull><FieldRef Name="User_x0020_Field" /></IsNotNull>
    </And>
</Where>
```