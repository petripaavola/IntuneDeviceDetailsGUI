# Intune Device Details GUI ver. 2.95 (updated 17.5.2022)
Go to script [IntuneDeviceDetailsGUI.ps1](./IntuneDeviceDetailsGUI.ps1)

**Version 2.95 is a huge update to the script's functionalities. Built-in search helps using this tool a lot.**

This Powershell based GUI/report helps Intune admins to see Intune device data in one view

Especially it shows what **Azure AD Groups and Intune filters** are used in Application and Configuration Assignments.

Assignment group information helps admins to understand why apps and configurations are targeted to devices and find possible bad assignments.

### GUI view
![IntuneDeviceDetailsGUI_v2.95.gif](https://github.com/petripaavola/IntuneDeviceDetailsGUI/blob/main/pics/IntuneDeviceDetailsGUI_v2.95.gif)

### Features
* **Search** with free keyword or use built-in Quick Filters
  * Keyword search with device name, serial, user email address, operating system, deviceId
  * searching with user email address also shows devices where user has logged-in (this is not shown in MEM/Intune search)
* **Show Application Assignments with AzureAD Groups and Filters information**
* **Show Configurations Assignments with AzureAD Groups and Filters information**
* Highlight assignment states with colors to easily see what is happening
  * For example highlight Not Applicable application assignment with yellow and usually you notice that Filter is reason for Not Applicable state
* Show Recently logged in user(s)
* Show OS Version support dates
* Show Device Group Memberships and membership rules
* Show Primary User Group Memberships and membership rules
* Show Latest Signed-In User Group Memberships and membership rules
* Intune JSON information on right side (this helps to understand what data there is inside Intune)
* **Hover on** Device, PrimaryUser, Latest logged-in user, Group, Application, Configuration, AssignmentGroup, Filter and many other places to get more information
  * There is lot of work done to get these (hover) ToolTips to show relevant information in easily readable format
* **Right click menus** to
  * Copy data to clipboard
  * Copy Dynamic Azure AD Group Membership rules to clipboard
  * Open specific resource in MEM/Intune web UI (device, autopilot device, user, group, application)
  * Copy Win32 Application custom detection and requirement scripts
* Script uses caching a lot. All used data is automatically downloaded once a day and during the day delta checks are done to Intune data
  * Idea is to use cached data but still every time double check that there are no changes in Intune
  * This saves Graph API bandwidth but still makes sure that data relevant (real time)
* Show Autopilot information
 * Applied Autopilot deployment profile
 * Applied Enrollment Status Page (this information is not available in MEM console)
 * Applied Enrollment Restrictions (this information is not available in MEM console)
* Support for Shared devices
  * Search with user email address and get devices where user has logged in
  * Latest logged in user information ToolTip (hover on)
  * Latest logged in user Azure AD Group memberships (and membeship rules)
  * Application and Configuration assignment are checked against latest logged in user's group memberships (if there is no PrimaryUser in device)


### Usage
**Prerequisities:**

**make sure you have installed Intune Powershell module and allow running Powershell scripts**
```
# Install Intune Powershell module
Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser

# Allow running Powershell scripts

# This allows running all signed scripts for current user
# This shows warning if script is not signed
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# This allows running all scripts for current user
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

**Running the script**
```
.\IntuneDeviceDetailsGUI.ps1
.\IntuneDeviceDetailsGUI.ps1 -Verbose

.\IntuneDeviceDetailsGUI.ps1 -Id 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
.\IntuneDeviceDetailsGUI.ps1 -IntuneDeviceId 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff

# Pipe Intune objects to script
Get-IntuneManagedDevice -Filter "devicename eq 'MyLoveMostPC'" | .\IntuneDeviceDetailsGUI.ps1

# Or create Device Management UI with Out-GridView
# Show devices in Out-GridView and for selected device show IntuneDeviceDetailsGUI
Get-IntuneManagedDevice | Out-GridView -OutputMode Single | .\IntuneDeviceDetailsGUI.ps1
```

### Future possible plans
* Never ever create Powershell UIs without using syncHash and multiple threads
  * This script was originally part of multithreaded tool so syncHash was not used inside this script and threading was done outside this script
  * Now this tool is self containing so in the next major version update syncHash and threading needs to be implemented to create responsive UI
* Integrate my other Application and Configuration Assignment reports into this toolset
* Create other views to show information about users, Azure AD Groups, Applications and Configurations
* Any other feature requests ?-)

## Disclaimer
This tool is provided "AS IS" without any warranties so please evaluate it in test environment before production use. It is provided as Powershell script so there is no closed code and you can evaluate everything it does. Trust is important when using Administrative user rights and tools in your production environment. I use this tool daily in production environments I manage myself.
