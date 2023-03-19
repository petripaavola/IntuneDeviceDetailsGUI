# Intune Device Details GUI ver. 2.974 Preview (updated 17.2.2023)
**This is also production version so we are up to date for now.**  
**Preview versions are development versions which I use daily with my production environments.**


Most likely reason for not being production version is that I would like to get feedback from testers that everything is ok in this version before making this version as a production release.


Download script [IntuneDeviceDetailsGUI_v2.974-Preview.ps1](./IntuneDeviceDetailsGUI_v2.974-Preview.ps1)

### Changelog v2.974
* Mouse cursor is changed to Waiting-mode (running circle) during report creation. This helps user understand that report creation process is running

### Changelog v2.973
* Fixed Configuration profiles report max 50 entries limit
  * Script originally fetched max 50 configuration profiles assignment intents
* Added **members count** to Assignment, Device and User groups
  * This helps understand **impact** of the deployment -> how many devices and/or users are targeted
* Added user selection dropdown for Application Assignments report (list of logged on users to device)
  * This is really important with **Shared devices** or if multiple users have logged on to device
* Updated Windows support dates
* Some other minor bug fixes

### Known bugs ###
* Intune New Store UWP Apps installState is not shown correctly on report
  * **Confirmed falsepositive**
  * With shared devices there is no user selected by default so that is the reason we don't get information for Store UWP App deployments.
  * Selecting user will show New Store UWP App Installation state for selected user
