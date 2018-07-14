# Sample demonstrates how to stop and start workflow
In some cases you may have stagnant workflows or need to update workflows that are in-flight.
You'll need to stop those WF's and start the new published workflow.
The following sample will create a special column.
Query the list for running instances.
Mark those List Items as stopped.
Stop the workflow instances.
Start the workflow subscription on preivously stopped instances.


```posh


Connect-SPIaC -url "https://[tenant].sharepoint.com/sites/siteA" -credentialname "spiqonline"

Stop-IaCWorkflowInstance -List "Application Requests" -View "Pending" -WorkflowName "Request Status" -WhatIf -Verbose

$cancelledWF = Stop-IaCWorkflowInstance -List "Application Requests" -View "Admin" -WorkflowName "Request Status" -Verbose

$listWF = Get-IaCWorkflowInstances -List "Application Requests" -WorkflowName "Request Status" -DeepScan -Verbose

$startedWF = Start-IaCWorkflowInstance -List "Application Requests" -View "Admin" -WorkflowName "Request Status" -Verbose

$startedWF | ConvertTo-Json -Depth 5 | Out-File "l:\temp\siterequestWF-started.json"


```