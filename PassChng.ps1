# Code for password change script goes here
Add-Type -AssemblyName System.Windows.Forms

# Check the current execution policy
$currentExecutionPolicy = Get-ExecutionPolicy

# Set the execution policy to Bypass temporarily
Set-ExecutionPolicy Bypass -Scope Process -Force

# Check if the ActiveDirectory module is already imported
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    # Try importing the ActiveDirectory module
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        # If importing fails, install the RSAT feature as admin
        Write-Host "Installing RSAT feature including Active Directory module..."

        # Define the command to install the RSAT feature
        $installCommand = "Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

        try {
            # Start a new PowerShell process as admin and execute the install command
            Start-Process powershell -ArgumentList "-NoProfile -Command `"$installCommand`"" -Verb RunAs -Wait
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "Failed to install RSAT feature: $errorMessage"
            return
        }

        # Import the ActiveDirectory module as user
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "Failed to import ActiveDirectory module: $errorMessage"
            return
        }
    }
}
# Create a form for password change
$passwordChangeForm = New-Object System.Windows.Forms.Form
$passwordChangeForm.Text = "AD Password Change"
$passwordChangeForm.Size = New-Object System.Drawing.Size(500, 500)
$passwordChangeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$passwordChangeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$passwordChangeForm.MaximizeBox = $false
$passwordChangeForm.BackColor = [System.Drawing.Color]::White

$logoPath = ".\Logo.png"  #REPLACE with the path to your logo image
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

# Event handler for the button click
$buttonChangePassword.Add_Click({
    # Code for password change script goes here
    $adusername = $textBoxUsername.Text
    $currentPassword = $textBoxCurrentPassword.Text | ConvertTo-SecureString -AsPlainText -Force
    $newPassword = $textBoxNewPassword.Text | ConvertTo-SecureString -AsPlainText -Force
    $reenteredNewPassword = $textBoxReenterNewPassword.Text | ConvertTo-SecureString -AsPlainText -Force

    # Check if any required text box is empty
    if ([string]::IsNullOrEmpty($adusername) -or [string]::IsNullOrEmpty($textBoxCurrentPassword.Text) -or [string]::IsNullOrEmpty($textBoxNewPassword.Text) -or [string]::IsNullOrEmpty($textBoxReenterNewPassword.Text)) {
        $message = "Fill in all the required details"
        Write-Host $message  # Displays error in PowerShell console
        [System.Windows.Forms.MessageBox]::Show($message, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Check if the new password meets the length requirement
    $newPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
    if ($newPasswordPlainText.Length -le 6) {
        $message = "Your new password must be more than 6 characters including letters, numbers, and symbols."
        [System.Windows.Forms.MessageBox]::Show($message, "Complexity Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Check if the entered passwords match
    $newPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
    $reenteredNewPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($reenteredNewPassword))
    if ($newPasswordPlainText -ne $reenteredNewPasswordPlainText) {
        $message = "The re-entered new password does not match the new password. Please try again."
        [System.Windows.Forms.MessageBox]::Show($message, "Mismatch Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    try {
        # Change the AD password
        Set-ADAccountPassword -Identity $adusername -OldPassword $currentPassword -NewPassword $newPassword

        # Convert the new password back to plain text
        $newPasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))

        $message = "Password has been changed successfully. Click OK to update it in the credential manager."
        [System.Windows.Forms.MessageBox]::Show($message, "SUCCESS", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Create a new PSCredential object with the updated password
        $targetName1 = "example1.com" #REPLACE (example1.com) with your Domain Name1
        $targetName2 = "*.example1.com" #REPLACE (*.example1.com) with your Domain Name1
        $targetName3 = "example2.com" #REPLACE (example2.com) with your Domain Name2
        $targetName4 = "*.example2.com" #REPLACE (*.example2.com) with your Domain Name2
        $username = "DOMAINNAME\$adusername" #REPLACE (DOMAINNAME) with your actual AD DoaminName
        $password = $newPasswordPlainText

        # Create the domain-specific or network address credential using vaultcmd
        $cmd1 = "vaultcmd /addcreds:`"Windows Credentials`" /credtype:`"Windows Domain Password Credential`" /identity:`"$username`" /authenticator:`"$password`" /resource:`"$targetName1`""
        $cmd2 = "vaultcmd /addcreds:`"Windows Credentials`" /credtype:`"Windows Domain Password Credential`" /identity:`"$username`" /authenticator:`"$password`" /resource:`"$targetName2`""
        $cmd3 = "vaultcmd /addcreds:`"Windows Credentials`" /credtype:`"Windows Domain Password Credential`" /identity:`"$username`" /authenticator:`"$password`" /resource:`"$targetName3`""
        $cmd4 = "vaultcmd /addcreds:`"Windows Credentials`" /credtype:`"Windows Domain Password Credential`" /identity:`"$username`" /authenticator:`"$password`" /resource:`"$targetName4`""

        $result1 = Invoke-Expression -Command $cmd1
        $result2 = Invoke-Expression -Command $cmd2
        $result3 = Invoke-Expression -Command $cmd3
        $result4 = Invoke-Expression -Command $cmd4

        $message = "The following have been successfully updated in the Credential Manager:`n"
        if ($result1 -eq "Credentials added successfully") { $message += "1. $targetName1`n" }
        if ($result2 -eq "Credentials added successfully") { $message += "2. $targetName2`n" }
        if ($result3 -eq "Credentials added successfully") { $message += "3. $targetName3`n" }
        if ($result4 -eq "Credentials added successfully") { $message += "4. $targetName4`n" }

        [System.Windows.Forms.MessageBox]::Show($message, "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Clear the plain text passwords from memory
        $currentPasswordPlainText, $newPasswordPlainText, $reenteredNewPasswordPlainText, $password = $null

        # Close the form
        $passwordChangeForm.Close()
    }
    catch {
        $errorMessage = $_.Exception.Message
        $message = "Password Change Error Details:`n$errorMessage"
        [System.Windows.Forms.MessageBox]::Show($message, "FAILED", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
        
        # Log the error
        Write-Host "Password Change Error: $errorMessage"
    }        
})

# Show the password change form
[void]$passwordChangeForm.ShowDialog()

# Restoring the original execution policy
Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force