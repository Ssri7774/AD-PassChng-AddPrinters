# UI-ADPassChng-AddPrinters
This repo helps to change the AD user's password and update it in the Windows Credential Manager. Also, includes to add the printers from print server for local users via local and external domain.

## New Features
- Added the field at the bottom of the UI to see which Domain Controller and Printer Server is in use
- Two scenarios to work this script form domain joined device (ForLocalDomains) and to run from external domain joined device (ForExtDomains)
- Username automatically entered based on logged in user (can be editable)
- Log file will be created in the directory of the script running from

## Download
Click [here](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/releases/download/latest/UI-ADPassChng-AddPrinters-main.zip) to download

## PREVIEWS

#### ADPassChng_AddPrinters: (Purpose - To change AD Password and Add Printers)

<img width="380" alt="image" src="https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/6225a338-0e8f-4ec5-8838-266c97dec5f9">

#### PassChng: (Purpose - To change AD Password)

<img width="385" alt="image" src="https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/e95e7420-4dae-4e35-8a30-935d1c4cff7e">

#### AddPrinters: (Purpose - To Add Printer(s) from the print server)

<img width="459" alt="image" src="https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/d26676bf-ecd9-49ab-8d48-0affa0d19b46">


## MODIFICATIONS

In the previous release, the modification have to made in the PowerShell Script. But now we have a JSON configuration file where the required details need to be updated to work this in your environment.

See the .json files in respective directories to modify the required fields. There are `comment` and `description` in the .json files which helps to understand what exact information need to be provided.


## How to use?

Just run the Setup.bat file to run the application which takes care of script execution policy to run as Bypass.

## Troubleshoot

### Potential Errors & Warnings

#### ForExtDomains

> To resolve this error, you need to close the application and run the Setup.bat again and enter the Correct Password.
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/4714d444-d3ae-4d90-a621-d20a431be94e)

> We get this prompt if the authentication with the print server (ps01.example1.com) failed. It fails if the password or username in the Credential Manager is wrong. So, you need to enter the current password in this prompt which will then update the Credential Manager with the currently logged-in username and the provided password and try again to authenticate with the print server.
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/87564741-1dde-4caf-83b3-c8a3f9f340cc)

#### For Local Domains

> To resolve this error, re-enter the correct password (no need to close the application).
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/5f0524c0-1201-40d7-b35f-e597cb0dbb28)

#### Common Errors

> If you click on the change password button without filling in any of the fields, it will cause this error. Make sure to fill in all the details before clicking on Change password.
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/6787e62c-3751-40ec-9e02-d8968db576f0)

> As mentioned in the warning, make sure that the complexity of the entered new password is more than 6 characters including letters, numbers, and symbols.
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/44f54d8e-d5ef-41b0-acb0-658ed5af0b9f)

> It means that the entered new password and re-entered new password are mismatched. Enter the correct password in both fields.
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/8c46cbef-c909-4d18-ac19-bc491ea2526a)

> This error can cause due to multiple reasons:
>
> - If the Password doesn't meet the Complexity criteria
>
> - If the limit of changing password per day is exceeded
>
> - If the user is trying to change the password same as previous passwords (It depends on the number of password history recognisation on your Domain Contoller)
>
>> ![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/4147742c-6f27-483a-a35b-ec64abbf9f73)

