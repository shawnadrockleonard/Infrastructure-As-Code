param($ProjectDir, $ConfigurationName, $TargetDir, $TargetFileName, $TargetName, $SolutionDir, $ProjectName)
begin 
{
	Write-Output ("Starting the post build script for {0}" -f $TargetDir)
}
process
{
	if($ConfigurationName -like "Debug*")
	{	
		$documentsFolder = [environment]::getfolderpath("mydocuments");

		Remove-Module -Name $ProjectName -ErrorAction SilentlyContinue
		$PSModuleHome = ("{0}\WindowsPowerShell\Modules\{1}" -f $documentsFolder,$TargetName)

		# Module folder there?
		if(Test-Path $PSModuleHome)
		{
			# Yes, empty it
			Remove-Item $PSModuleHome\* -Force -Recurse
		} 
		else 
		{
			# No, create it
			New-Item -Path $PSModuleHome -ItemType Directory -Force >$null # Suppress output
		}

		Write-Host "Copying files from $TargetDir to $PSModuleHome"
		Copy-Item "$TargetDir\*.dll" -Destination "$PSModuleHome"
		Copy-Item "$TargetDir\*help.xml" -Destination "$PSModuleHome"
		Copy-Item "$TargetDir\samplelog.config" -Destination  "$PSModuleHome" -ErrorAction SilentlyContinue
		Copy-Item "$TargetDir\ModuleFiles\*.psd1" -Destination  "$PSModuleHome"
		Copy-Item "$TargetDir\ModuleFiles\*.ps1xml" -Destination "$PSModuleHome"
		Copy-Item "$TargetDir\ModuleFiles\*.resx" -Destination "$PSModuleHome"
	} 
	elseif ($ConfigurationName -like "Release*")
	{
		$distDir = "$SolutionDir\dist";

		# Dist folder there? If so, empty it.
		if(Test-Path $distDir)
		{
			Remove-Item $distDir\*
		} else {
			# Create folder
			New-Item -Path "$distDir" -ItemType Directory -Force >$null # Suppress output
		}
		# Copy files to 'dist' folder
		Write-Host "Copying files from $TargetDir to $distDir"
		Copy-Item "$TargetDir\*.dll" -Destination "$distDir"
		Copy-Item "$TargetDir\*help.xml" -Destination "$distDir"
		Copy-Item "$TargetDir\samplelog.config" -Destination  "$distDir" -ErrorAction SilentlyContinue
		Copy-Item "$TargetDir\ModuleFiles\*.psd1" -Destination  "$distDir"
		Copy-Item "$TargetDir\ModuleFiles\*.ps1xml" -Destination "$distDir"
		Copy-Item "$SolutionDir\install.ps1" -Destination "$distDir"
		ii $distDir
	}
}
end
{
	Write-Output ("Finish the post build script for {0}" -f $TargetDir)
}
	