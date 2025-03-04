# Define your GPO names
$GPO1 = "Group Policy 1" ##The GPO we're replacing
$GPO2 = "Group Policy 2" ##The GPO being added in

# Define log file path
$logFile = "C:\GPO_Link_Changes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" ##Update this to the path you want to output the logs to

# Import the Group Policy module
Import-Module GroupPolicy

# Get all linked locations for GPO1
$linkedLocations = Get-GPInheritance -Target "DomainNameHere" | ForEach-Object { 
    $_.GpoLinks | Where-Object { $_.DisplayName -eq $GPO1 } | ForEach-Object { $_.Target }
}

# Remove duplicates if any
$linkedLocations = $linkedLocations | Sort-Object -Unique

if ($linkedLocations.Count -eq 0) {
    Write-Host "No linked locations found for $GPO1."
    exit
}

Write-Host "Found $($linkedLocations.Count) locations linked to $GPO1. Processing..."

# Initialize log data
$logData = @()

foreach ($location in $linkedLocations) {
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
