# Import the Active Directory module
Import-Module ActiveDirectory

# Specify the user you want to remove from extensionAttribute6
$userToRemove = Read-Host "Enter the username to remove from extensionAttribute6"

# Function to get groups from a text file or manual input
function Get-TargetGroups {
    Write-Host "`nHow would you like to specify the groups?"
    Write-Host "1. Enter group names manually"
    Write-Host "2. Import from groups.txt in current directory"
    $choice = Read-Host "Enter your choice (1 or 2)"

    $groupNames = @()
    
    if ($choice -eq "1") {
        Write-Host "`nEnter group names (press Enter after each group, type 'done' when finished):"
        do {
            $input = Read-Host
            if ($input -ne "done") {
                $groupNames += $input
            }
        } while ($input -ne "done")
    }
    elseif ($choice -eq "2") {
        # Get and display the current directory
        $currentPath = Get-Location
        Write-Host "`nCurrent Directory: $currentPath"
        
        # Get and display the script directory
        $scriptPath = $PSScriptRoot
        if (!$scriptPath) {
            $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        Write-Host "Script Directory: $scriptPath"
        
        $filePath = Join-Path $scriptPath "groups.txt"
        Write-Host "Looking for groups.txt at: $filePath"
        
        if (Test-Path $filePath) {
            $groupNames = Get-Content $filePath
            Write-Host "Successfully loaded groups from groups.txt"
            Write-Host "Groups found in file:"
            $groupNames | ForEach-Object { Write-Host "- $_" }
        }
        else {
            Write-Host "groups.txt not found at: $filePath" -ForegroundColor Red
            Write-Host "`nPlease ensure groups.txt exists in the same directory as the script." -ForegroundColor Yellow
            Write-Host "Current files in script directory:" -ForegroundColor Yellow
            Get-ChildItem $scriptPath | Select-Object Name
            exit
        }
    }
    else {
        Write-Host "Invalid choice. Please run the script again." -ForegroundColor Red
        exit
    }
    
    return $groupNames
}

# Function to properly format owner list
function Format-OwnerList {
    param (
        [string]$currentValue,
        [string]$userToRemove
    )
    
    # Split the current value by comma and trim each entry
    $owners = $currentValue -split ',' | ForEach-Object { $_.Trim() }
    
    # Remove the specified user (case-insensitive)
    $remainingOwners = $owners | Where-Object { $_ -ne $userToRemove }
    
    # Join the remaining owners with comma and space
    $newValue = $remainingOwners -join ', '
    
    return $newValue
}

# Get the target groups
$targetGroupNames = Get-TargetGroups

# Create a log file in the same directory as the script
$scriptPath = $PSScriptRoot
if (!$scriptPath) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$logFile = Join-Path $scriptPath "attribute_changes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to log file
function Write-ToLog {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Add-Content -Path $logFile
}

Write-ToLog "Starting script execution"

# Get the specified groups that have the user in extensionAttribute6
$groups = @()
$notFoundGroups = @()
$noAttributeGroups = @()

foreach ($groupName in $targetGroupNames) {
    try {
        $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Properties extensionAttribute6
        if ($group) {
            if ($group.extensionAttribute6 -like "*$userToRemove*") {
                $groups += $group
            }
            else {
                $noAttributeGroups += $groupName
            }
        }
        else {
            $notFoundGroups += $groupName
        }
    }
    catch {
        Write-ToLog "Error finding group ${groupName} - $($_.Exception.Message)"
        $notFoundGroups += $groupName
    }
}

# Report status
Write-Host "`nStatus Report:"
Write-Host "Found $($groups.Count) groups containing user '$userToRemove' in extensionAttribute6"
if ($notFoundGroups.Count -gt 0) {
    Write-Host "Groups not found: $($notFoundGroups -join ', ')" -ForegroundColor Yellow
}
if ($noAttributeGroups.Count -gt 0) {
    Write-Host "Groups without user in extensionAttribute6: $($noAttributeGroups -join ', ')" -ForegroundColor Yellow
}

Write-ToLog "Found $($groups.Count) groups containing user '$userToRemove'"
Write-ToLog "Groups not found: $($notFoundGroups -join ', ')"
Write-ToLog "Groups without user in extensionAttribute6: $($noAttributeGroups -join ', ')"

if ($groups.Count -eq 0) {
    Write-Host "No valid groups found for processing. Exiting script."
    Write-ToLog "No valid groups found for processing"
    exit
}

# Show current values and preview changes
foreach ($group in $groups) {
    $currentValue = $group.extensionAttribute6
    $newValue = Format-OwnerList -currentValue $currentValue -userToRemove $userToRemove
    
    Write-Host "`nGroup: $($group.Name)"
    Write-Host "Current Value: $currentValue"
    Write-Host "New Value: $newValue"
    
    Write-ToLog "Group $($group.Name)"
    Write-ToLog "Current Value: $currentValue"
    Write-ToLog "Proposed New Value: $newValue"
}

# Prompt for confirmation
$confirmation = Read-Host "`nDo you want to proceed with these changes? (Y/N)"

if ($confirmation -eq 'Y') {
    Write-Host "`nProcessing changes..."
    Write-ToLog "User confirmed changes - beginning updates"
    
    foreach ($group in $groups) {
        try {
            $currentValue = $group.extensionAttribute6
            $newValue = Format-OwnerList -currentValue $currentValue -userToRemove $userToRemove
            
            # If newValue is empty, clear the attribute
            if ([string]::IsNullOrWhiteSpace($newValue)) {
                Set-ADGroup -Identity $group.DistinguishedName -Clear extensionAttribute6
                Write-Host "Successfully cleared extensionAttribute6 for group $($group.Name)"
                Write-ToLog "Successfully cleared extensionAttribute6 for group $($group.Name)"
            }
            else {
                Set-ADGroup -Identity $group.DistinguishedName -Replace @{extensionAttribute6 = $newValue}
                Write-Host "Successfully updated group $($group.Name)"
                Write-ToLog "Successfully updated group $($group.Name) - New value: $newValue"
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            $errorDetails = $_.Exception.GetType().FullName
            Write-Host "Error updating group $($group.Name): $errorMessage" -ForegroundColor Red
            Write-Host "Error Type: $errorDetails" -ForegroundColor Red
            Write-ToLog "ERROR updating group $($group.Name) - $errorMessage"
            Write-ToLog "ERROR Type: $errorDetails"
        }
    }
    
    Write-Host "`nChanges completed. Check $logFile for detailed log."
    Write-ToLog "Script execution completed"
}
else {
    Write-Host "Operation cancelled by user."
    Write-ToLog "Operation cancelled by user"
}