# InfrastructureAsCode.PowerShell #
The InfrastructureAsCode binary Powershell module can be run during debug or built and deployed to extend and reuse in command-line or PowerShell scripts.

From documentation a module is [Understanding a Windows PowerShell module]https://msdn.microsoft.com/en-us/library/dd878324(v=vs.85).aspx
A module is a set of related Windows PowerShell functionalities, grouped together as a convenient unit (usually saved in a single directory). By defining a set of related script files, assemblies, and related resources as a module, you can reference, load, persist, and share your code much easier than you would otherwise.
The main purpose of a module is to allow the modularization (ie, reuse and abstraction) of Windows PowerShell code. For example, the most basic way of creating a module is to simply save a Windows PowerShell script as a .psm1 file. Doing so allows you to control (ie, make public or private) the functions and variables contained in the script. Saving the script as a .psm1 file also allows you to control the scope of certain variables. Finally, you can also use cmdlets such as Install-Module to organize, install, and use your script as building blocks for larger solutions.

Why did we choose a binary module and not a script module?  https://msdn.microsoft.com/en-us/library/dd878342(v=vs.85).aspx With a binary module we can reuse the InfrastructureAsCode.CORE project, leverage multi-threading, perform numerous activities while leveraging open source managed code projects to effect continous and constant change to a SharePoint environment.  Many examples you'll find online use a console application.  With a binary module we can stitch together console applications as CmdLets in a single project and reusue core functionality.

# How to run and debug
If you attempt to debug the project and it does not launch Powershell, navigate to the Project Debug tab and specified the external program powershell.exe with a -noexit argument.   Take a look at the screenshot below for more detail.  
You should set the start external project to PowerShell.exe or Powershell_ISE.exe if you prefer the Integrated Scripting Environment.   Powershell ISE allow you to write specific scripts and include additional binary modules with ease.

<img src="https://raw.githubusercontent.com/pinch-perfect/Infrastructure-As-Code/master/InfrastructureAsCode.Powershell/imgs/project-config-powershell-debug.PNG" />
<caption>User project specific settings in the project properties</caption>

# Examples
Once you are able to debug the InfrastructureAsCode.PowerShell module you'll be presented with a command-line or ISE command-line.  At the command line the first step before you can run any Cmd-Let is to Connect to your SharePoint instance.  This version supports ADFS authentication (thanks to the PnP-Powershell project), inline credentials, and hashed credentials.  

The following command will connect to your tenant with the specified UserName.  You will be prompted for a password 1 time.  The module will hash the password and store it locally.  The next time you run this command it will use the hashed secure string and connect without a prompt.   You can use this command to automate a schedule task or to run a specific command without being prompted for a password each time.
```powershell
Connect-SPIaC –Url https://[tenant].sharepoint.com –UserName "[user]@[tenant].onmicrosoft.com
```


The following command will connect to your tenant with the Credential prompt which will ask for a username and password.  You could also prompt for credentials and store it in a variable and pass the variable to the parameter
```powershell
Connect-SPIaC –Url https://[tenant].sharepoint.com -Credentials (Get-Credential)
```
or
```powershell
$credentials = (Get-Credential)
Connect-SPIaC –Url https://[tenant].sharepoint.com -Credentials $credentials
```

After you have successfull connected to the the SharePoint URL, received your ClientContext, and perform a variety of commands you should disconnect.  This frees the ClientContext from memory and is good practice.
```powershell
Disconnect-SPIaC
```


# Documentation
Please note: not all Cmd-Lets are well documented.  This will come over time with command examples for how you should run a particular command.  We will provide a number of code samples that include references to these individual Cmd-Lets.



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