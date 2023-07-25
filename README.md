# UI-ADPassChng-AddPrinters
This repo helps to change the AD user's password and update it in the Windows Credential Manager. Also, includes to add the printers from print server

## PREVIEWS

#### ADPassChng_AddPrinters: (Purpose - To change AD Password and Add Printers)

![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/4583b41c-81f6-4f5a-b77c-687906bf0600)

#### PassChng: (Purpose - To change AD Password)

![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/40af7fda-4997-4c30-bd58-d0534a8349f2)

#### AddPrinters: (Purpose - To Add Printer(s) from the print server)

![image](https://github.com/Ssri7774/UI-ADPassChng-AddPrinters/assets/95307763/9b907983-314d-43cf-9e17-d9304d86e9ba)


## MODIFICATIONS

Note: You need to replace the following places from the script to use it for your environment (Tip: Use 'Find' feature to see the places to 'REPLACE').

### ADPassChng_AddPrinters.ps1
1. Line 15 - #REPLACE (.\Logo.png) with the path to your logo image (Optional)
2. Line 225 - #REPLACE (example1.com) with your Domain Name1
3. Line 226 - #REPLACE (*.example.com) with your Domain Name1
4. Line 227 - #REPLACE (example2.com) with your Domain Name2
5. Line 228 - #REPLACE (*.example2.com) with your Domain Name2
6. Line 229 - #REPLACE (DOMAINNAME) with your actual AD DoaminName
7. Line 366 - #REPLACE (ps01.example.com) with your print server Name with FQDN

### PassChng.ps1
1. Line 53 - #REPLACE (.\Logo.png) with the path to your logo image (Optional)
2. Line 171 - #REPLACE (example1.com) with your Domain Name1
3. Line 172 - #REPLACE (*.example.com) with your Domain Name1
4. Line 173 - #REPLACE (example2.com) with your Domain Name2
5. Line 174 - #REPLACE (*.example2.com) with your Domain Name2
6. Line 175 - #REPLACE (DOMAINNAME) with your actual AD DoaminName

### AddPrinters.ps1
1. Line 13 - #REPLACE (.\Logo.png) with the path to your logo image (Optional)
2. Line 87 - #REPLACE (ps01.example.com) with your print server Name with FQDN

In ADPassChng_AddPrinters.ps1 & PassChng.ps1, use example1.com & example2.com if you have mutiple domains in the forest (This applies to the environment where the organisation changed the name from example1 to example2 and still using example1 in the backend). This part only adds all the possible user credentials to the Windows Credential Manager.
