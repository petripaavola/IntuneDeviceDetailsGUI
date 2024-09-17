# Intune Device Details GUI ver. 3.0 (updated 2024-09-17)

### Note! Script uses now Microsoft.Graph.Authentication module ###

**Go script [IntuneDeviceDetailsGUI.ps1](./IntuneDeviceDetailsGUI.ps1)**

**Version 3.0 shows Remediation scripts.**
* Shows Remediation scripts in third assignment category
* Hovering on top of Remediation script you can see for example output (hover over status column) and schedule (hover over AssignmentGroup)

**Version 2.985 update to Microsoft Graph module.**
* Added Graph API scope **DeviceLocalCredential.Read.All** to get LAPS passwords

**Version 2.982 update to Microsoft Graph module.**
* Added Graph API scope **DeviceManagementServiceConfig.Read.All** to get deviceEnrollmentConfiguration information
* Changed Bitlocker Recovery Keys and LAPS Password **dateTime to current timezone**
* Improved Graph authentication so script will fail if we don't see TenantId (which should mean authentication failed somehow even if it may have seem to be worked ok)

**Version 2.98 update to Microsoft Graph module.**
* **Script uses now Microsoft.Graph.Authentication Powershell module**
  * You can install Microsoft Graph module with command: **Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser**
  * Script uses only Read Scopes for Graph API
* New feature to show **Bitlocker Recovery Keys**
* New feature to show **Windows LAPS Password**
* Device group memberships shows also nested groups memberships which will help show Application and Configuration target groups

**Version 2.974 is another bigger update.**

*  **There was bug in Configuration Profiles processing so make sure to update to this newer version (it limited to 50 policies).**
*  You can now see **impact** of Assignments meaning you can see number of users and/or devices affected by Assignment.  
*  You can also select logged on user to show Application assignments to.  
*  Something small but also really big on it's own way. Mouse cursor changes to busy while searching :)  

**Version 2.95 is a huge update to the script's functionalities. Built-in search helps using this tool a lot.**

This Windows Powershell based GUI/report helps Intune admins to see Intune device data in one view

Especially it shows what **Azure AD Groups and Intune filters** are used in Application and Configuration Assignments.

Assignment group information helps admins to understand why apps and configurations are targeted to devices and find possible bad assignments.

### GUI view
![IntuneDeviceDetailsGUI_v2.95.gif](https://github.com/petripaavola/IntuneDeviceDetailsGUI/blob/main/pics/IntuneDeviceDetailsGUI_v2.95.gif)

### Features
* **Search** with free keyword or use built-in Quick Filters
  * Keyword search with device name, serial, user email address, operating system, deviceId
  * searching with user email address also shows devices where user has logged-in (this is not shown in MEM/Intune search)
* **Show Application Assignments with EntraID Groups and Filters information**
* **Show Configurations Assignments with EntraID Groups and Filters information**
* **Show Remediation script Assignments with EntraID Groups and Filters information**
* **Show Bitlocker Recovery Keys**
* **Show Windows LAPS Password**
* Highlight assignment states with colors to easily see what is happening
  * For example highlight Not Applicable application assignment with yellow and usually you notice that Filter is reason for Not Applicable state
* Show Recently logged in user(s)
* Show OS Version support dates
* Show Device Group Memberships and membership rules
* Show Primary User Group Memberships and membership rules
* Show Latest Signed-In User Group Memberships and membership rules
* Intune JSON information on right side (this helps to understand what data there is inside Intune)
* **Hover on** to get ToolTip on Device, PrimaryUser, Latest logged-in user, Group, Application, Configuration, AssignmentGroup, Filter and many other places to get more information
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
  * Applied Enrollment Status Page
  * Applied Enrollment Restrictions
  * Autopilot Deployment Profile JSON
  * Autopilot device JSON
* Support for Shared devices
  * Search with user email address and get devices where user has logged in
  * Latest logged in user information ToolTip (hover on)
  * Latest logged in user Azure AD Group memberships (and membeship rules)
  * Application and Configuration assignment are checked against latest logged in user's group memberships (if there is no PrimaryUser in device)


### Usage
**Prerequisities:**

Run script in **Windows Powershell**. Windows Presentations Framework (WPF) based GUIs don´t work with Powershell core.

**Make sure you have installed Microsoft.Graph.Authentication module and allow running Powershell scripts**
```
# Install Microsoft.Graph.Authentication module
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser

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

### Intune user permissions needed to run this script
* Script uses user login so login with user who has at least **read permissions** to Intune and permissions to use Intune Powershell module
* Script does only GET operations and 2 POST operations (to get Intune report) so at this time this is read only tool
* For most restrictive permissions this script has been tested even with **Intune Read Only Operators role** and script works with that role also

### Future possible plans
* Never ever create Powershell UIs without using syncHash and multiple threads
  * This script was originally part of multithreaded tool so syncHash was not used inside this script and threading was done outside this script
  * Now this tool is self containing so in the next major version update syncHash and threading needs to be implemented to create responsive UI
* Integrate my other Application and Configuration Assignment reports into this toolset
* Create other views to show information about users, Azure AD Groups, Applications and Configurations
* Any other feature requests ?-)

## Disclaimer
This tool is provided "AS IS" without any warranties so please evaluate it in test environment before production use. It is provided as Powershell script so there is no closed code and you can evaluate everything it does. Trust is important when using Administrative user rights and tools in your production environment. I use this tool daily in production environments I manage myself.
