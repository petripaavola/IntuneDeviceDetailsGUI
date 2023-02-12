# Intune Device Details GUI ver. 2.973 Preview (updated 12.2.2023)
**Preview versions are development versions which I use daily with my production environments.**


Most likely reason for not being production version is that I would like to get feedback from testers that everything is ok in this version before making this version as a production release.


Download script [IntuneDeviceDetailsGUI_v2.973-Preview.ps1](./IntuneDeviceDetailsGUI_v2.973-Preview.ps1)

### Changelog
* Fixed Configuration profiles report max 50 entries limit
  * Script originally fetched max 50 configuration profiles assignment intents
* Added **members count** to Assignment, Device and User groups
  * This helps understand **impact** of the deployment -> how many devices and/or users are targeted
* Added user selection dropdown for Application Assignments report (list of logged on users to device)
  * This is really important with **Shared devices** or if multiple users have logged on to device
* Updated Windows support dates
* Some other minor bug fixes
