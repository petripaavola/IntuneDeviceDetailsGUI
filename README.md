# Intune Device Details GUI ver. 2.3
Go to script [IntuneDeviceDetailsGUI.ps1](./IntuneDeviceDetailsGUI.ps1)

This Powershell based GUI/report helps Intune admins to see Intune device data in one view

Especially it shows what **Azure AD Groups and what Intune filters** are used in Application and Configuration assignments.

Assignment group information helps admins to understand why apps and configurations are targeted to devices and find possible bad assignments.

### Other information also shown
* Recently logged in user(s)
* OS Version support dates
* Device Group Memberships
* Primary User Group Memberships (hover on top of Primary User)
* Autopilot deployment profile JSON (hover on top of Autopilot profile)
* Intune device JSON information on right side (this helps to understand what data there is inside Intune)

### GUI view
![IntuneDeviceDetailsGUI.png](https://www.petripaavola.fi/IntuneDeviceDetailsGUI.png)

### Usage
* make sure you have installed Intune Powershell module (**Install-Module -Name Microsoft.Graph.Intune**)

```
.\IntuneDeviceDetailsGUI.ps1
.\IntuneDeviceDetailsGUI.ps1 -deviceName MyLoveMostPC
.\IntuneDeviceDetailsGUI.ps1 -Id 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
.\IntuneDeviceDetailsGUI.ps1 -IntuneDeviceId 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
.\IntuneDeviceDetailsGUI.ps1 -serialNumber 1234567890

# Pipe Intune objects to script
Get-IntuneManagedDevice -Filter "devicename eq 'MyLoveMostPC'" | .\IntuneDeviceDetailsGUI.ps1
'MyLoveMostPC' | .\IntuneDeviceDetailsGUI.ps1

# Or create Device Management UI with Out-GridView
# Show devices in Out-GridView and for selected device show IntuneDeviceDetailsGUI
Get-IntuneManagedDevice | Out-GridView -OutputMode Single | .\IntuneDeviceDetailsGUI.ps1
```

## Disclaimer
This tool is provided "AS IS" without any warranties so please evaluate it in test environment before production use. It is provided as Powershell script so there is no closed code and you can evaluate everything it does. Trust is important when using Administrative user rights and tools in your production environment. I use this tool daily in production environments I manage myself.
