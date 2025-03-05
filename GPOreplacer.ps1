# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to retrieve all GPOs in the domain
function Get-AllGPOs {
    Import-Module GroupPolicy
    return (Get-GPO -All | Select-Object -ExpandProperty DisplayName | Sort-Object)
}

# Function to retrieve domain name dynamically
function Get-DomainName {
    try {
        return (Get-ADDomain).DNSRoot
    } catch {
        return "Enter Manually"
    }
}

# Function to create a dropdown UI
function Get-UserSelection {
    param (
        [string]$Prompt,
        [array]$Options
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Input"
    $form.Size = New-Object System.Drawing.Size(350, 180)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Size = New-Object System.Drawing.Size(320, 20)
    $label.Location = New-Object System.Drawing.Point(10, 10)

    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Size = New-Object System.Drawing.Size(320, 20)
    $comboBox.Location = New-Object System.Drawing.Point(10, 40)
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $comboBox.Items.AddRange($Options)

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Size = New-Object System.Drawing.Size(100, 30)
    $buttonOK.Location = New-Object System.Drawing.Point(75, 80)
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $buttonOK

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
    $buttonCancel.Location = New-Object System.Drawing.Point(180, 80)
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $buttonCancel

    $form.Controls.Add($label)
    $form.Controls.Add($comboBox)
    $form.Controls.Add($buttonOK)
    $form.Controls.Add($buttonCancel)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $comboBox.Text
    } else {
        Write-Host "User cancelled input. Exiting script."
        exit
    }
}

# Function to create a checkbox UI for selecting locations
function Get-UserSelectionMulti {
    param (
        [string]$Prompt,
        [array]$Options
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Prompt
    $form.Size = New-Object System.Drawing.Size(400, 500)
    $form.StartPosition = "CenterScreen"

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Size = New-Object System.Drawing.Size(370, 350)
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
    $checkedListBox.Items.AddRange($Options)

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Size = New-Object System.Drawing.Size(100, 30)
    $buttonOK.Location = New-Object System.Drawing.Point(75, 400)
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $buttonOK

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
    $buttonCancel.Location = New-Object System.Drawing.Point(200, 400)
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $buttonCancel

    $form.Controls.Add($checkedListBox)
    $form.Controls.Add($buttonOK)
    $form.Controls.Add($buttonCancel)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $checkedListBox.CheckedItems
    } else {
        Write-Host "User cancelled selection. Exiting script."
        exit
    }
}

# Function to display progress bar
function Show-Progress {
    param (
        [int]$PercentComplete
    )
    Write-Progress -Activity "Processing GPO changes..." -PercentComplete $PercentComplete
}

# Fetch available GPOs and domain name
$gpoList = Get-AllGPOs
$domainList = @(Get-DomainName, "Enter Manually")

# Prompt user for GPO selections and domain
$GPO1 = Get-UserSelection -Prompt "Select the GPO to replace (GPO1):" -Options $gpoList
$GPO2 = Get-UserSelection -Prompt "Select the new GPO to add (GPO2):" -Options $gpoList
$Domain = Get-UserSelection -Prompt "Select the domain name:" -Options $domainList

# Define log file path
$logFile = "C:\temp\GPO_Link_Changes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Import the Group Policy module
Import-Module GroupPolicy

# Get all linked locations for GPO1
$linkedLocations = Get-GPInheritance -Domain $Domain -All | 
    Where-Object { $_.GpoLinks -match "DisplayName=$GPO1" } | 
    Select-Object -ExpandProperty Target

# Remove duplicates if any
$linkedLocations = $linkedLocations | Sort-Object -Unique

if ($linkedLocations.Count -eq 0) {
    Write-Host "No linked locations found for $GPO1."
    exit
}

Write-Host "Found $($linkedLocations.Count) locations linked to $GPO1."

# Allow user to select locations to modify
$selectedLocations = Get-UserSelectionMulti -Prompt "Select locations to update" -Options $linkedLocations

# Initialize log data
$logData = @()

# Process selected locations
$totalLocations = $selectedLocations.Count
$counter = 0

foreach ($location in $selectedLocations) {
    $counter++
    Show-Progress -PercentComplete (($counter / $totalLocations) * 100)

    # Log entry for the location
    $entry = @{
        "Timestamp" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "Location"  = $location
        "GPO1_Removed" = $false
        "GPO2_Linked" = $false
    }

    try {
        # Add GPO2 to the same location
        Write-Host "Linking $GPO2 to $location..."
        New-GPLink -Name $GPO2 -Target $location -Confirm:$false
        $entry["GPO2_Linked"] = $true
    } catch {
        Write-Host "Failed to link $GPO2 to $location. Error: $_" -ForegroundColor Red
    }

    try {
        # Remove GPO1 from the location
        Write-Host "Removing $GPO1 from $location..."
        Remove-GPLink -Name $GPO1 -Target $location -Confirm:$false
        $entry["GPO1_Removed"] = $true
    } catch {
        Write-Host "Failed to remove $GPO1 from $location. Error: $_" -ForegroundColor Red
    }

    # Add entry to log data
    $logData += New-Object PSObject -Property $entry
}

# Export log data to CSV
$logData | Export-Csv -Path $logFile -NoTypeInformation
Write-Host "Log saved to $logFile"

Write-Host "Completed updating GPO links."
