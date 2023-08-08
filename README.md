# UI-ADPassChng-AddPrinters
This repo helps to change the AD user's password and update it in the Windows Credential Manager. Also, includes to add the printers from print server

# New Features
- Added the field at the bottom of the UI to see which Domain Controller and Printer Server is in use
- Two scenarios to work this script form domain joined device (ForLocalDomains) and to run from external domain joined device (ForExtDomains)
- Username automatically entered based on logged in user (can we editable)
- Log file will be created in the directory of the script running from

## PREVIEWS

#### ADPassChng_AddPrinters: (Purpose - To change AD Password and Add Printers)

![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/46dcd4ae-5473-4a90-b478-f0adaeea1ab5)

#### PassChng: (Purpose - To change AD Password)

<img width="384" alt="image" src="https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/b05ead5d-1f0b-4589-8eda-772e2d52df3c">

#### AddPrinters: (Purpose - To Add Printer(s) from the print server)

<img width="459" alt="image" src="https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/d26676bf-ecd9-49ab-8d48-0affa0d19b46">


## MODIFICATIONS

In the previous release, the modification have to made in the PowerShell Script. But now we have a JSON configuration file where the required details need to be updated to work this in you environment.

See the .json files in respective directories to modify the required fields. There are `comment` and `description` in the .json files which helps to understand what exact information need to be provided.
