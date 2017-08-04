<#
    .SYNOPSIS
		Executes a series of cmdlets to create a JSON file from a class definition
		Will then use the JSON to provision the resources in SharePoint
    .PARAMETER RelativeOrFullPath
		Specifies the relative path to the JSON file to be used in the site configuration.
    .OUTPUTS
		Nothing

	Example
		From your home 'Documents' directory
		cd ("{0}\\{1}\\Documents" -f $env:HOMEDRIVE, $env:HOMEPATH)
		$RelativeOrFullPath = Full Path to your Project Folder c:\[YOUR REPO FOLDER]\development-samples\Sample02\AppFiles\

	.\WindowsPowerShell\Modules\InfrastructureAsCode.Powershell.Sample01\script-configure-sample01.ps1 -RelativeOrFullPath "C:\Users\sleonard.CONTOSO\Documents\SampleDefinitions"
#>  
[CmdletBinding(HelpURI='http://aka.ms/pinch-perfect')]
Param(
    [Parameter(Mandatory = $true)]
    [String]$RelativeOrFullPath
)
BEGIN 
{
	# Configure context to SharePoint site
	# Connect-SPIaC -Url "https://[tenant].sharepoint.com" -UserName "[user]@[tenant].onmicrosoft.com"
}
PROCESS
{
	try {

		Select-IaCSampleQuery -Verbose
	}
	catch {
		Write-Error $_.Exception[0]
	}
	finally {
		Disconnect-SPIaC
	}
}