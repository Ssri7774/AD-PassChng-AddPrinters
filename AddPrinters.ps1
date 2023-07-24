# Create a form for printer installation
$printerInstallationForm = New-Object System.Windows.Forms.Form
$printerInstallationForm.Text = "Printer Selection"
$printerInstallationForm.Size = New-Object System.Drawing.Size(600, 555)
$printerInstallationForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$printerInstallationForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$printerInstallationForm.MaximizeBox = $false
$printerInstallationForm.BackColor = [System.Drawing.Color]::White

$logoPath = ".\Logo.png"  # REPLACE with the path to your logo image
if (Test-Path $logoPath) {
    $printerInstallationLogo = New-Object System.Windows.Forms.PictureBox
    $printerInstallationLogo.Image = [System.Drawing.Image]::FromFile($logoPath)
    $printerInstallationLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $printerInstallationLogo.Location = New-Object System.Drawing.Point(20, 350)
    $printerInstallationLogo.Size = New-Object System.Drawing.Size(545, 170)
    $printerInstallationForm.Controls.Add($printerInstallationLogo)
}

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

# Create checkboxes for printer selection
$checkBoxes = @()
try {
    $printers = Get-Printer -ComputerName ps01.example.com | Where-Object { $_.Shared -eq $true } | Select-Object Name,ComputerName | Sort-Object Name #REPLACE (ps01.example.com) with your print server Name with FQDN
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Host "Failed to get printers: $errorMessage"
    [System.Windows.Forms.MessageBox]::Show("Failed to get printers. Error details:`n$errorMessage", "Printer Retrieval Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

$printersHashTable = @{}

$yPosition = 30
$checkBoxWidth = 500

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

# Event handler for the button click
$buttonInstallPrinter.Add_Click({
    # Code for printer installation goes here
    # Get the selected printers
    $selectedPrinters = $checkBoxes | Where-Object { $_.Checked }
    Write-Host "Adding $printerUNCPath printer. . . . ."

    # Check if any printer is selected
    if ($selectedPrinters.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one printer.", "Printer Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Set the maximum value of the progress bar
    $progressBar.Maximum = $selectedPrinters.Count

    # Variable to track if any errors occurred
    $errorOccurred = $false

    # Loop through the selected printers
    foreach ($printer in $selectedPrinters) {
        $printerName = $printer.Text
        $printerUNCPath = $printersHashTable[$printerName]

        # Output the constructed UNC path for debugging
        Write-Host "Printer UNC Path: $printerUNCPath"
        
        # Update the progress bar
        $progressBar.Value++

        try {
            # Attempt to add the printer using the UNC path
            Add-Printer -ConnectionName $printerUNCPath -ErrorAction Stop
        } catch {
            $errorMessage = $_.Exception.Message
            Write-Host $errorMessage  # Displays error in PowerShell console
            [System.Windows.Forms.MessageBox]::Show("Failed to install printer '$printerName'. Error details:`n$errorMessage", "Printer Installation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $errorOccurred = $true
            break
        }
    }

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

# Show the printer installation form
[void]$printerInstallationForm.ShowDialog()