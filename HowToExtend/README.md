# InfrastructureAsCode.PowerShell.HowToExtend
InfrastructureAsCode.PowerShell.HowToExtend is a sample PowerShell project which attempt to assist you conceptually with how to extend PnP-Powershell or InfrastructureAsCode.  We will gladly accept your pull requests if and when you share common infrastructure capabilities.  If you want to leverage a solid base with yours; you can follow this project.    You should be able to sync this locally and follow the guidance found in https://github.com/pinch-perfect/Infrastructure-As-Code/tree/master/InfrastructureAsCode.Powershell to ensure you can debug or build the project.   

# Running the sample
This project will compile as a Binary Powershell module and deploy to your [Users/Documents/WindowsPowershell/Module/InfrastructureAsCode.Powershell.HowToExtend] folder.  If you attempt to debug the project and it does not launch Powershell, navigate to the Project Debug tab and specify the external program powershell.exe with a -noexit argument.  
You can however, build the project and then open a powershell window and execute the Cmd-Lets as demonstrated:
```powershell
Connect-SPIaC –Url https://[tenant].sharepoint.com –UserName "[user]@[tenant].onmicrosoft.com

Select-IaCSampleQuery -Verbose #Simple verbose write out to the command-line

Disconnect-SPIaC
```


# Video demonstration
Here is a video to watch how you can go from Build to configuration of a SharePoint site.

<img src="https://raw.githubusercontent.com/pinch-perfect/Infrastructure-As-Code/master/HowToExtend/imgs/build-and-deploy.png" />
https://mix.office.com/embed/1hqxzxbj2h51c


# Sample Scripts
At the root of the project are sample .ps1 files which demonstrate a scripted Cmd-Let running the binary module commands.

Open script-configure-assets.ps1 and you'll see the following:
<img src="https://raw.githubusercontent.com/pinch-perfect/Infrastructure-As-Code/master/HowToExtend/imgs/project-sample-script.PNG" />

Open script-configure-provision.ps1 and you'll see the following:
<img src="https://raw.githubusercontent.com/pinch-perfect/Infrastructure-As-Code/master/HowToExtend/imgs/project-sample-script-provisioning-resources.PNG" />


### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

Microsoft provides programming examples for illustration only, without 
warranty either expressed or implied, including, but not limited to, the
implied warranties of merchantability and/or fitness for a particular 
purpose.  

This sample assumes that you are familiar with the programming language
being demonstrated and the tools used to create and debug procedures. 
Microsoft support professionals can help explain the functionality of a
particular procedure, but they will not modify these examples to provide
added functionality or construct procedures to meet your specific needs. 
If you have limited programming experience, you may want to contact a 
Microsoft Certified Partner or the Microsoft fee-based consulting line 
at (800) 936-5200. 

It is your responsiblity to review the code and understand what is being performed.