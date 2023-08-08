Start-Transcript -Path $PSScriptRoot\Logs.log -Append

$Configuration = Get-Content $PSScriptRoot\\confExtDomain.json | ConvertFrom-Json

function Get-BestDomainController {
    param (
        $DomainControllers 
    )
    ForEach($DomainController in $DomainControllers) {
        if(Test-Connection -ComputerName $DomainController.Name -Count 1 -ErrorAction SilentlyContinue) {
            return $DomainController
        }
    }
}

$BestDomainController = Get-BestDomainController -DomainControllers $Configuration.Curriculum.DomainController

# function Get-BestDomainName {
#     param (
#         $Domains 
#     )
#     ForEach($Domain in $Domains) {
#         if(Test-Connection $Domain.Name -Count 1 -ErrorAction SilentlyContinue) {
#             return $Domain
#         }
#     }
# }

# $BestDomainName = Get-BestDomainName -Domains $Configuration.Curriculum.Domain

function Get-BestPrintServer {
    param (
        $PrintServers 
    )
    ForEach($PrintServer in $PrintServers) {
        if(Test-Connection -ComputerName $PrintServer.Name -Count 1 -ErrorAction SilentlyContinue) {
            return $PrintServer
        }
    }
}

$BestPrintServer = Get-BestPrintServer -PrintServers $Configuration.Curriculum.PrintServer

# Define Logo
$logoPath = $Configuration.Curriculum.Media.Logo

# Define Authentication Domin Name a.k.a NetBIOS
$netBIOS = $Configuration.Curriculum.NetBIOS.Name

# Get the full username with domain name using whoami
$UpdatedAs = whoami

# Extract only the username from the full username
$user = $UpdatedAs -replace '.*\\'

#Username with NetBIOS
$fullUsername = "$netBIOS\$user"

# This function used to update the creds in Credential manager
function Update-CredentialManager {
    param (
        [string]$username,
        [string]$auth,
        [string]$updatedBy,
        [string]$targetName
    )

    # Create the command for updating the Credential Manager using vaultcmd
    $cmd = ('vaultcmd /addcreds:"Windows Credentials" /credtype:"Windows Domain Password Credential" /identity:{0} /authenticator:{1} /resource:{2} /savedBy:{3}' -f $username, '$auth', $targetName, $updatedBy)

    # Execute the command
    Invoke-Expression -Command $cmd
}

Add-Type -AssemblyName System.Windows.Forms

# Create the main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Select Action"
$mainForm.Size = New-Object System.Drawing.Size(500, 500)
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.MaximizeBox = $false
# Set the background color
$mainForm.BackColor = [System.Drawing.Color]::White
# $mainForm.BackColor = [System.Drawing.Color]::FromArgb(95, 146, 155) # Use this for RGB colours

# Add Logo
$mainLogo = New-Object System.Windows.Forms.PictureBox
$mainLogo.Image = [System.Drawing.Image]::FromFile($logoPath)
$mainLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$mainLogo.Location = New-Object System.Drawing.Point(20, 220)
$mainLogo.Size = New-Object System.Drawing.Size(445, 250)
$mainForm.Controls.Add($mainLogo)

# Create a label for the heading text
$labelHeading = New-Object System.Windows.Forms.Label
$labelHeading.Text = "Choose as per your requirement"
$labelHeading.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
$labelHeading.AutoSize = $true
$labelHeading.Location = New-Object System.Drawing.Point(75, 30)
$mainForm.Controls.Add($labelHeading)

# Create a button for password change
$buttonPasswordChange = New-Object System.Windows.Forms.Button
$buttonPasswordChange.Text = "Password Change"
$buttonPasswordChange.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$buttonPasswordChange.Location = New-Object System.Drawing.Point(50, 100)
$buttonPasswordChange.Size = New-Object System.Drawing.Size(120, 80)
$buttonPasswordChange.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$mainForm.Controls.Add($buttonPasswordChange)

# Create a button for printer installation
$buttonPrinterInstallation = New-Object System.Windows.Forms.Button
$buttonPrinterInstallation.Text = "Add `nPrinter(s)"
$buttonPrinterInstallation.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$buttonPrinterInstallation.Location = New-Object System.Drawing.Point(315, 100)
$buttonPrinterInstallation.Size = New-Object System.Drawing.Size(120, 80)
$buttonPrinterInstallation.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$mainForm.Controls.Add($buttonPrinterInstallation)

# Event handler for the password change button click
$buttonPasswordChange.Add_Click({
    # Hide the main form
    $mainForm.Hide()

    # Shows the DC using from the mentioned DCs in json file
    Write-Host "Using DC: $($BestDomainController.Name)"
    
    # Create a "Please wait" form
    $Main2PCwaitForm = New-Object System.Windows.Forms.Form
    $Main2PCwaitForm.Text = "In Progress"
    $Main2PCwaitForm.Size = New-Object System.Drawing.Size(400, 100)
    $Main2PCwaitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $Main2PCwaitForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Main2PCwaitForm.MaximizeBox = $false
    $Main2PCwaitForm.BackColor = [System.Drawing.Color]::White

    # Create a label for the "Please wait" message
    $Main2PClabelWaitMessage = New-Object System.Windows.Forms.Label
    $Main2PClabelWaitMessage.Text = "Please wait. Connecting to LDAP Server. . ."
    $Main2PClabelWaitMessage.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
    $Main2PClabelWaitMessage.AutoSize = $true
    $Main2PClabelWaitMessage.Location = New-Object System.Drawing.Point(10, 10)
    $Main2PCwaitForm.Controls.Add($Main2PClabelWaitMessage)

    # Show the "Please wait" form as modeless
    $Main2PCwaitForm.Show()

    # Create a form for password change
    $passwordChangeForm = New-Object System.Windows.Forms.Form
    $passwordChangeForm.Text = "AD Password Change"
    $passwordChangeForm.Size = New-Object System.Drawing.Size(500, 510)
    $passwordChangeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $passwordChangeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $passwordChangeForm.MaximizeBox = $false
    $passwordChangeForm.BackColor = [System.Drawing.Color]::White

    if (Test-Path $logoPath) {
        $passwordChangeLogo = New-Object System.Windows.Forms.PictureBox
        $passwordChangeLogo.Image = [System.Drawing.Image]::FromFile($logoPath)
        $passwordChangeLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $passwordChangeLogo.Location = New-Object System.Drawing.Point(20, 280)
        $passwordChangeLogo.Size = New-Object System.Drawing.Size(445, 170)
        $passwordChangeForm.Controls.Add($passwordChangeLogo)
    }

    # Create labels, textboxes, and buttons for password change
    $labelUsername = New-Object System.Windows.Forms.Label
    $labelUsername.Location = New-Object System.Drawing.Point(20, 25)
    $labelUsername.Size = New-Object System.Drawing.Size(175, 20)
    $labelUsername.Text = "AD Username:"
    $labelUsername.Font = New-Object System.Drawing.Font("Arial", 11)
    $passwordChangeForm.Controls.Add($labelUsername)

    $textBoxUsername = New-Object System.Windows.Forms.TextBox
    $textBoxUsername.Location = New-Object System.Drawing.Point(195, 20)
    $textBoxUsername.Size = New-Object System.Drawing.Size(280, 30)
    $textBoxUsername.Font = New-Object System.Drawing.Font("Arial", 12)
    $textBoxUsername.Text = $user
    $passwordChangeForm.Controls.Add($textBoxUsername)

    $labelCurrentPassword = New-Object System.Windows.Forms.Label
    $labelCurrentPassword.Location = New-Object System.Drawing.Point(20, 75)
    $labelCurrentPassword.Size = New-Object System.Drawing.Size(175, 20)
    $labelCurrentPassword.Text = "Current Password:"
    $labelCurrentPassword.Font = New-Object System.Drawing.Font("Arial", 11)
    $passwordChangeForm.Controls.Add($labelCurrentPassword)

    $textBoxCurrentPassword = New-Object System.Windows.Forms.TextBox
    $textBoxCurrentPassword.Location = New-Object System.Drawing.Point(195, 70)
    $textBoxCurrentPassword.Size = New-Object System.Drawing.Size(280, 30)
    $textBoxCurrentPassword.PasswordChar = '*'
    $textBoxCurrentPassword.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordChangeForm.Controls.Add($textBoxCurrentPassword)

    $labelNewPassword = New-Object System.Windows.Forms.Label
    $labelNewPassword.Location = New-Object System.Drawing.Point(20, 125)
    $labelNewPassword.Size = New-Object System.Drawing.Size(175, 20)
    $labelNewPassword.Text = "New Password:"
    $labelNewPassword.Font = New-Object System.Drawing.Font("Arial", 11)
    $passwordChangeForm.Controls.Add($labelNewPassword)

    $textBoxNewPassword = New-Object System.Windows.Forms.TextBox
    $textBoxNewPassword.Location = New-Object System.Drawing.Point(195, 120)
    $textBoxNewPassword.Size = New-Object System.Drawing.Size(280, 30)
    $textBoxNewPassword.PasswordChar = '*'
    $textBoxNewPassword.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordChangeForm.Controls.Add($textBoxNewPassword)

    $labelReenterNewPassword = New-Object System.Windows.Forms.Label
    $labelReenterNewPassword.Location = New-Object System.Drawing.Point(20, 175)
    $labelReenterNewPassword.Size = New-Object System.Drawing.Size(175, 30)
    $labelReenterNewPassword.Text = "Re-enter New Password:"
    $labelReenterNewPassword.Font = New-Object System.Drawing.Font("Arial", 11)
    $passwordChangeForm.Controls.Add($labelReenterNewPassword)

    $textBoxReenterNewPassword = New-Object System.Windows.Forms.TextBox
    $textBoxReenterNewPassword.Location = New-Object System.Drawing.Point(195, 170)
    $textBoxReenterNewPassword.Size = New-Object System.Drawing.Size(280, 30)
    $textBoxReenterNewPassword.PasswordChar = '*'
    $textBoxReenterNewPassword.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordChangeForm.Controls.Add($textBoxReenterNewPassword)

    $buttonChangePassword = New-Object System.Windows.Forms.Button
    $buttonChangePassword.Location = New-Object System.Drawing.Point(150, 230)
    $buttonChangePassword.Size = New-Object System.Drawing.Size(185, 40)
    $buttonChangePassword.Text = "Change Password"
    $buttonChangePassword.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $buttonChangePassword.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
    $passwordChangeForm.Controls.Add($buttonChangePassword)

    $labelDCName = New-Object System.Windows.Forms.Label
    $labelDCName.Text = "Using DC: $($BestDomainController.Name)"
    $labelDCName.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
    $labelDCName.AutoSize = $true
    $labelDCName.Location = New-Object System.Drawing.Point(10, 450)
    $passwordChangeForm.Controls.Add($labelDCName)

    # Close the "Please wait" form as modeless
    $Main2PCwaitForm.Close()

    # Event handler for the button click
    $buttonChangePassword.Add_Click({

        # Create a "Please wait" form
        $CPwaitForm = New-Object System.Windows.Forms.Form
        $CPwaitForm.Text = "In Progress"
        $CPwaitForm.Size = New-Object System.Drawing.Size(400, 100)
        $CPwaitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $CPwaitForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $CPwaitForm.MaximizeBox = $false
        $CPwaitForm.BackColor = [System.Drawing.Color]::White

        # Create a label for the "Please wait" message
        $CPlabelWaitMessage = New-Object System.Windows.Forms.Label
        $CPlabelWaitMessage.Text = "Please wait while changing the password..."
        $CPlabelWaitMessage.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
        $CPlabelWaitMessage.AutoSize = $true
        $CPlabelWaitMessage.Location = New-Object System.Drawing.Point(20, 20)
        $CPwaitForm.Controls.Add($CPlabelWaitMessage)

        # Code for password change script goes here
        $enteredUsername = $textBoxUsername.Text
        $currentPassword = $textBoxCurrentPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $newPassword = $textBoxNewPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $reenteredNewPassword = $textBoxReenterNewPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $dcserver = $BestDomainController.Name
        $enteredFullUsername = "$netBIOS\$enteredUsername"
        # Show the "Please wait" form as modeless
        $CPwaitForm.Show()

        $currentPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($currentPassword))
        # Check if any required text box is empty
        if ([string]::IsNullOrEmpty($textBoxUsername.Text) -or [string]::IsNullOrEmpty($textBoxCurrentPassword.Text) -or [string]::IsNullOrEmpty($textBoxNewPassword.Text) -or [string]::IsNullOrEmpty($textBoxReenterNewPassword.Text)) {
            $message = "Fill in all the required details"
            Write-Host $message  # Display error in PowerShell console
            # Close the "Please wait" form
            $CPwaitForm.Close()
            [System.Windows.Forms.MessageBox]::Show($message, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check if the new password meets the length requirement
        $newPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
        Write-Host $newPasswordSecure
        if ($newPasswordPlainText.Length -le 6) {
            $message = "Your new password must be more than 6 characters including letters, numbers, and symbols."
            # Close the "Please wait" form
            $CPwaitForm.Close()
            [System.Windows.Forms.MessageBox]::Show($message, "Complexity Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Check if the entered passwords match
        $newPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
        $reenteredNewPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($reenteredNewPassword))
        if ($newPasswordPlainText -ne $reenteredNewPasswordPlainText) {
            $message = "The re-entered new password does not match the new password. Please try again."
            # Close the "Please wait" form
            $CPwaitForm.Close()
            [System.Windows.Forms.MessageBox]::Show($message, "Mismatch Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            # Create an array to store the target names
            $preTargetNames = @()
    
            foreach ($preDomain in $Configuration.Curriculum.Domain) {
                $preTargetName = $preDomain.Name
                $preTargetNames += $preTargetName
    
                # Call the function with the correct password
                Update-CredentialManager -username $enteredFullUsername -auth $currentPasswordPlainText -updatedBy $UpdatedAs -targetName $preTargetName
            }
            Write-Host "Updated Credential Manager with provided Current Password to authenticate with ADSI: `n-user: $enteredFullUsername `n-password: Current [Cannot be displayed] `n-targetName: $preTargetNames `n-updatedBy: $UpdatedAs"
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "Failed to authenticate with ADSI. Entered Username/Password is incorrect: $errorMessage"
            # Close the "Please wait" form
            $CPwaitForm.Close()
            [System.Windows.Forms.MessageBox]::Show("Failed to authenticate with ADSI. Entered Username/Password is incorrect `nError details:`n$errorMessage", "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }

        #Authentication with ADSI (Active Directory Services Interfaces)
        [ADSI]$useraccount="WinNT://$($dcserver)/$enteredUsername"
               
        if ($useraccount -ne $null) {
            # Attempt to change the password using Kerberos authentication
            $success = $false

            try {
                $useraccount.ChangePassword($currentPasswordPlainText, $newPasswordPlainText)
                $success = $true
                # Write-Host "Password changed successfully."
                Write-Host "$enteredFullUsername - Password has been changed successfully" -ForegroundColor Green
                $message = "Password has been changed successfully. Click OK to update it in the credential manager."
                # Close the "Please wait" form
                $CPwaitForm.Close()
                [System.Windows.Forms.MessageBox]::Show($message, "SUCCESS", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                $errorMessage = $_.Exception.Message
                $message = "Password Change Error Details:`n$errorMessage"
                # Close the "Please wait" form
                $CPwaitForm.Close()
                [System.Windows.Forms.MessageBox]::Show($message, "FAILED", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
                # Log the error
                Write-Host "Password Change Error: $errorMessage"
            }

            # Add Credentials to the Windoes Credentials Manager
            if ($success) {            
                # Create a "Please wait" form
                $CMwaitForm = New-Object System.Windows.Forms.Form
                $CMwaitForm.Text = "In Progress"
                $CMwaitForm.Size = New-Object System.Drawing.Size(400, 100)
                $CMwaitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $CMwaitForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
                $CMwaitForm.MaximizeBox = $false
                $CMwaitForm.BackColor = [System.Drawing.Color]::White

                # Create a label for the "Please wait" message
                $CMlabelWaitMessage = New-Object System.Windows.Forms.Label
                $CMlabelWaitMessage.Text = "Please wait. Adding credentials to Credential Manager..."
                $CMlabelWaitMessage.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
                $CMlabelWaitMessage.AutoSize = $true
                $CMlabelWaitMessage.Location = New-Object System.Drawing.Point(20, 20)
                $CMwaitForm.Controls.Add($CMlabelWaitMessage)

                # Show the "Please wait" form as modeless
                $CMwaitForm.Show()

                # Create an array to store the target names
                $TargetNames = @()

                foreach ($Domain in $Configuration.Curriculum.Domain) {
                    $TargetName = $Domain.Name
                    $TargetNames += $TargetName

                    # Call the function with the correct password
                    Update-CredentialManager -username $enteredFullUsername -auth $enteredPasswordPlainText -updatedBy $UpdatedAs -targetName $TargetName
                }
                Write-Host "Updated Credential Manager with provided NEW-CREDs: `n-user: $enteredFullUsername `n-password: New [Cannot be displayed] `n-targetName: $TargetNames `n-updatedBy: $UpdatedAs"

                Write-Host "$enteredFullUsername - Updated Credential Manager with provided NEW-CREDs" -ForegroundColor Green
                $message = "Updated Credential Manager with provided NEW-CREDs"
                # Close the "Please wait" form
                $CMwaitForm.Close()
                [System.Windows.Forms.MessageBox]::Show($message, "Creds Update", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
                # Close the form
                $passwordChangeForm.Close()
            }
        }
    })

    $passwordChangeForm.Add_FormClosing({
        # Show the main form again
        $mainForm.Show()
    })

    # Show the password change form
    [void]$passwordChangeForm.ShowDialog()
})

# Event handler for the printer installation button click
$buttonPrinterInstallation.Add_Click({
    # Hide the main form
    $mainForm.Hide()

    # Shows Print Server from the mentioned print servers in json file
    Write-Host "Using PS: $($BestPrintServer.Name)"

    function Get-CredentialsAndSave {
        # Create a form for the password prompt
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = "Enter Current Password"
        $passwordForm.Size = New-Object System.Drawing.Size(350, 150)
        $passwordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $passwordForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $passwordForm.MaximizeBox = $false
        $passwordForm.BackColor = [System.Drawing.Color]::White

        # Create a label for the password prompt
        $labelPasswordPrompt = New-Object System.Windows.Forms.Label
        $labelPasswordPrompt.Text = "Please enter your current password:"
        $labelPasswordPrompt.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
        $labelPasswordPrompt.AutoSize = $true
        $labelPasswordPrompt.Location = New-Object System.Drawing.Point(10, 10)
        $passwordForm.Controls.Add($labelPasswordPrompt)

        # Create a TextBox for entering the password
        $textBoxPassword = New-Object System.Windows.Forms.TextBox
        $textBoxPassword.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
        $textBoxPassword.Location = New-Object System.Drawing.Point(10, 40)
        $textBoxPassword.Size = New-Object System.Drawing.Size(320, 24)
        $textBoxPassword.PasswordChar = "*" # Mask the password
        $passwordForm.Controls.Add($textBoxPassword)

        # Create a button for confirming the password
        $buttonConfirmPassword = New-Object System.Windows.Forms.Button
        $buttonConfirmPassword.Text = "OK"
        $buttonConfirmPassword.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $buttonConfirmPassword.Location = New-Object System.Drawing.Point(10, 80)
        $buttonConfirmPassword.Size = New-Object System.Drawing.Size(100, 30)
        $buttonConfirmPassword.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
        $passwordForm.Controls.Add($buttonConfirmPassword)

        # Create an event handler for the button click
        $buttonConfirmPassword.Add_Click({
            # Get the entered password
            $enteredPassword = $textBoxPassword.Text | ConvertTo-SecureString -AsPlainText -Force
            $passwordForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $enteredPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($enteredPassword))

            # Create an array to store the target names
            $updTargetNames = @()

            foreach ($updDomain in $Configuration.Curriculum.Domain) {
                $updTargetName = $updDomain.Name
                $updTargetNames += $updTargetName

                # Call the function with the correct password
                Update-CredentialManager -username $fullUsername -auth $enteredPasswordPlainText -updatedBy $UpdatedAs -targetName $updTargetName
            }
            Write-Host "Updated Credential Manager with PROVIDED password to authenticate with Print Server: `n-user: $fullUsername `n-password: Provided [Cannot be displayed] `n-targetName: $updTargetNames `n-updatedBy: $UpdatedAs"
        })

        # Show the password prompt form and wait for user input
        $passwordForm.ShowDialog()
    }

    # Create a "Please wait" form
    $Main2APwaitForm = New-Object System.Windows.Forms.Form
    $Main2APwaitForm.Text = "In Progress"
    $Main2APwaitForm.Size = New-Object System.Drawing.Size(400, 100)
    $Main2APwaitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $Main2APwaitForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Main2APwaitForm.MaximizeBox = $false
    $Main2APwaitForm.BackColor = [System.Drawing.Color]::White

    # Create a label for the "Please wait" message
    $Main2APlabelWaitMessage = New-Object System.Windows.Forms.Label
    $Main2APlabelWaitMessage.Text = "Please wait. Pulling the Printers list..."
    $Main2APlabelWaitMessage.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
    $Main2APlabelWaitMessage.AutoSize = $true
    $Main2APlabelWaitMessage.Location = New-Object System.Drawing.Point(10, 10)
    $Main2APwaitForm.Controls.Add($Main2APlabelWaitMessage)

    # Show the "Please wait" form as modeless
    $Main2APwaitForm.Show()

    # Create a form for printer installation
    $printerInstallationForm = New-Object System.Windows.Forms.Form
    $printerInstallationForm.Text = "Printer Selection"
    $printerInstallationForm.Size = New-Object System.Drawing.Size(600, 580)
    $printerInstallationForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $printerInstallationForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $printerInstallationForm.MaximizeBox = $false
	$printerInstallationForm.BackColor = [System.Drawing.Color]::White

    if (Test-Path $logoPath) {
        $printerInstallationLogo = New-Object System.Windows.Forms.PictureBox
        $printerInstallationLogo.Image = [System.Drawing.Image]::FromFile($logoPath)
        $printerInstallationLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $printerInstallationLogo.Location = New-Object System.Drawing.Point(20, 350)
        $printerInstallationLogo.Size = New-Object System.Drawing.Size(545, 170)
        $printerInstallationForm.Controls.Add($printerInstallationLogo)
    }

    $labelPSName = New-Object System.Windows.Forms.Label
    $labelPSName.Text = "Using PS: $($BestPrintServer.Name)"
    $labelPSName.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
    $labelPSName.AutoSize = $true
    $labelPSName.Location = New-Object System.Drawing.Point(10, 520)
    $printerInstallationForm.Controls.Add($labelPSName)

    # Create a label for the heading text
    $labelHeading = New-Object System.Windows.Forms.Label
    $labelHeading.Text = "Please choose one or more printer(s) from below to install or click Exit."
    $labelHeading.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $labelHeading.AutoSize = $true
    $labelHeading.Location = New-Object System.Drawing.Point(5, 5)
    $printerInstallationForm.Controls.Add($labelHeading)

    # Create a label for the search box
    $labelSearch = New-Object System.Windows.Forms.Label
    $labelSearch.Text = "Search Printers:"
    $labelSearch.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $labelSearch.AutoSize = $true
    $labelSearch.Location = New-Object System.Drawing.Point(5, 30)
    $printerInstallationForm.Controls.Add($labelSearch)

    # Create a TextBox for the search box
    $textBoxSearch = New-Object System.Windows.Forms.TextBox
    $textBoxSearch.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $textBoxSearch.Location = New-Object System.Drawing.Point(130, 30)
    $textBoxSearch.Size = New-Object System.Drawing.Size(415, 24)
    $printerInstallationForm.Controls.Add($textBoxSearch)

    # Event handler for the search box text change
    $textBoxSearch.Add_TextChanged({
        $searchText = $textBoxSearch.Text.Trim().ToLower()

        # Sort the checkboxes based on the search text
        $sortedCheckBoxes = $checkBoxes | Sort-Object { $_.Text.ToLower().IndexOf($searchText) }

        foreach ($checkBox in $sortedCheckBoxes | Sort-Object) {
            if ($checkBox.Text.ToLower().Contains($searchText)) {
                $checkBox.Visible = $true
            } else {
                $checkBox.Visible = $false
            }
        }

        # Reorder the visible checkboxes so that searched printer(s) appear at the top
        $visibleCheckBoxes = $sortedCheckBoxes | Where-Object { $_.Visible }
        $yPosition = 30
        foreach ($checkBox in $visibleCheckBoxes | Sort-Object) {
            $checkBox.Location = New-Object System.Drawing.Point(20, $yPosition)
            $yPosition += 30
        }
    })

    # Create a panel control with a vertical scrollbar
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 55)
    $panel.Size = New-Object System.Drawing.Size(565, 220)
    $panel.AutoScroll = $true  # Enable vertical scrollbar
    $printerInstallationForm.Controls.Add($panel)

    # Create a group box for printer selection
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Select Printers"
    $groupBox.Location = New-Object System.Drawing.Point(0, 0)
    $groupBox.AutoSize = $true
    $panel.Controls.Add($groupBox)
    
    try {
        $printers = Get-Printer -ComputerName $BestPrintServer.Name | Where-Object { $_.Shared -eq $true } | Select-Object Name,ComputerName | Sort-Object Name

         # If the error is related to invalid credentials, prompt the user for new credentials
         if (!$printers) {
            Write-Host "Trying again to authenticate with PROVIDED password"
            Get-CredentialsAndSave
            # Retry the Get-Printer command
            try {
                $printers = Get-Printer -ComputerName $BestPrintServer.Name | Where-Object { $_.Shared -eq $true } | Select-Object Name,ComputerName | Sort-Object Name
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Host "Failed to get printers: $errorMessage"
                [System.Windows.Forms.MessageBox]::Show("Provided password may be wrong. Failed to get printers. `nError details:`n$errorMessage", "Printer Retrieval Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Unable to communicate with Print Server: `n$errorMessage"
        [System.Windows.Forms.MessageBox]::Show("Cannot connect to Print Server. `nError details:`n$errorMessage", "Printer Retrieval Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

    }

    # Create checkboxes for printer selection
    $checkBoxes = @()

    $printersHashTable = @{}  # Initialize the hashtable here

    $yPosition = 30
    $checkBoxWidth = 500  # Adjust the width of the checkboxes

    foreach ($printer in $printers) {
        $printerName = $printer.Name
        $printerServer = $printer.ComputerName
        $printerUNCPath = "\\" + $printerServer + "\" + $printer.Name
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = $printerName
        $checkBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
        $checkBox.Location = New-Object System.Drawing.Point(20, $yPosition)
        $checkBox.Size = New-Object System.Drawing.Size($checkBoxWidth, 20)
        $groupBox.Controls.Add($checkBox)
        $checkBoxes += $checkBox

        # Store the UNC path in a hashtable with the printer name as the key
        $printersHashTable[$printerName] = $printerUNCPath

        $yPosition += 30
    }

    # Create a button for printer installation
    $buttonInstallPrinter = New-Object System.Windows.Forms.Button
    $buttonInstallPrinter.Text = "Install"
    $buttonInstallPrinter.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $buttonInstallPrinter.Location = New-Object System.Drawing.Point(150, 285)
    $buttonInstallPrinter.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
    $printerInstallationForm.Controls.Add($buttonInstallPrinter)

    # Create an "Exit" button
    $buttonExit = New-Object System.Windows.Forms.Button
    $buttonExit.Text = "Exit"
    $buttonExit.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Regular)
    $buttonExit.Location = New-Object System.Drawing.Point(360, 285)
    $buttonExit.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
    $printerInstallationForm.Controls.Add($buttonExit)

    # Create a progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 320)
    $progressBar.Size = New-Object System.Drawing.Size(545, 30)
    $printerInstallationForm.Controls.Add($progressBar)

    # Close the "Please wait" form as modeless
    $Main2APwaitForm.Close()

    # Event handler for the button click
    $buttonInstallPrinter.Add_Click({
        # Code for printer installation goes here
        # Get the selected printers
        $selectedPrinters = $checkBoxes | Where-Object { $_.Checked }
    
        # Check if any printer is selected
        if ($selectedPrinters.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one printer.", "Printer Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Create a "Please wait" form
        $APwaitForm = New-Object System.Windows.Forms.Form
        $APwaitForm.Text = "In Progress"
        $APwaitForm.Size = New-Object System.Drawing.Size(400, 100)
        $APwaitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $APwaitForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $APwaitForm.MaximizeBox = $false
        $APwaitForm.BackColor = [System.Drawing.Color]::White

        # Create a label for the "Please wait" message
        $APlabelWaitMessage = New-Object System.Windows.Forms.Label
        $APlabelWaitMessage.Text = "Please wait. Adding the printer(s)..."
        $APlabelWaitMessage.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
        $APlabelWaitMessage.AutoSize = $true
        $APlabelWaitMessage.Location = New-Object System.Drawing.Point(10, 10)
        $APwaitForm.Controls.Add($APlabelWaitMessage)

        # Show the "Please wait" form as modeless
        $APwaitForm.Show()
    
        # Set the maximum value of the progress bar
        $progressBar.Maximum = $selectedPrinters.Count
    
        # Variable to track if any errors occurred
        $errorOccurred = $false
    
        # Loop through the selected printers
        foreach ($printer in $selectedPrinters) {
            $printerName = $printer.Text
            $printerUNCPath = $printersHashTable[$printerName]
            
            # Update the progress bar
            $progressBar.Value++

            try {
                # Attempt to add the printer using the UNC path
                Add-Printer -ConnectionName $printerUNCPath -ErrorAction Stop
                # Output the constructed UNC path for debugging
                Write-Host "Installed: $printerUNCPath"
            } catch {
                $errorMessage = $_.Exception.Message
                Write-Host $errorMessage  # Display error in PowerShell console
                [System.Windows.Forms.MessageBox]::Show("Failed to install printer '$printerUNCPath'. Error details:`n$errorMessage", "Printer Installation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $errorOccurred = $true
                break
            }
        }

        # Close the "Please wait" form
        $APwaitForm.Close()
    
        if ($errorOccurred) {
            # Display error message if any error occurred
            [System.Windows.Forms.MessageBox]::Show("Printer installation failed.", "Printer Installation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } else {
            # Display success message if no errors occurred
            [System.Windows.Forms.MessageBox]::Show("Printer(s) added successfully.", "Printer Installation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    
        # Reset the progress bar
        $progressBar.Value = 0
    
        # Close the printer installation form
        $printerInstallationForm.Close()
    })

    # Event handler for the exit button click
    $buttonExit.Add_Click({
		# Prompt the user if they want to exit
        $result = [System.Windows.Forms.MessageBox]::Show("Do you want to exit without installing the printer(s)?", "Printer Selection", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
			# Close the printer installation form
			$printerInstallationForm.Close()
		}
    })

    # Event handler for the form closing event
    $printerInstallationForm.Add_FormClosing({
        # Show the main form again
        $mainForm.Show()
    })    

    # Show the printer installation form
    [void]$printerInstallationForm.ShowDialog()
})

# Clear the plain text passwords from memory
$currentPasswordPlainText, $newPasswordPlainText, $reenteredNewPasswordPlainText, $auth, $currentPassword = $null

# Show the main form
[void]$mainForm.ShowDialog()

Stop-Transcript