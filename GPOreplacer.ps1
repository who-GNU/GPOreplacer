# Define your GPO names
$GPO1 = "Group Policy 1"
$GPO2 = "Group Policy 2"

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

foreach ($location in $linkedLocations) {
    # Add GPO2 to the same location
    Write-Host "Linking $GPO2 to $location..."
    New-GPLink -Name $GPO2 -Target $location -Confirm:$false

    # Remove GPO1 from the location
    Write-Host "Removing $GPO1 from $location..."
    Remove-GPLink -Name $GPO1 -Target $location -Confirm:$false
}

Write-Host "Completed updating GPO links."
