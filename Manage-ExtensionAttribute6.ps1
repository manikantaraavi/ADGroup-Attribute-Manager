# Import the Active Directory module
Import-Module ActiveDirectory

# Function to display menu and get user choice
function Show-Menu {
    Clear-Host
    Write-Host "================ AD ExtensionAttribute6 Manager ================"
    Write-Host "1: Remove user from extensionAttribute6"
    Write-Host "2: Add user to extensionAttribute6"
    Write-Host "3: Edit existing users in extensionAttribute6"
    Write-Host "4: Exit"
    Write-Host "=========================================================="
    
    $choice = Read-Host "Please enter your choice (1-4)"
    return $choice
}

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
        $scriptPath = $PSScriptRoot
        if (!$scriptPath) {
            $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        $filePath = Join-Path $scriptPath "groups.txt"
        
        if (Test-Path $filePath) {
            $groupNames = Get-Content $filePath
            Write-Host "Successfully loaded groups from groups.txt"
            Write-Host "Groups found in file:"
            $groupNames | ForEach-Object { Write-Host "- $_" }
        }
        else {
            Write-Host "groups.txt not found at: $filePath" -ForegroundColor Red
            Write-Host "`nPlease ensure groups.txt exists in the same directory as the script." -ForegroundColor Yellow
            exit
        }
    }
    return $groupNames
}

# Function to properly format owner list
function Format-OwnerList {
    param (
        [string]$currentValue,
        [string]$userToAdd
    )
    
    if ([string]::IsNullOrWhiteSpace($currentValue)) {
        return $userToAdd
    }
    
    $owners = $currentValue -split ',' | ForEach-Object { $_.Trim() }
    if ($owners -contains $userToAdd) {
        return $currentValue
    }
    
    return "$currentValue, $userToAdd"
}

# Function to remove user from owner list
function Remove-FromOwnerList {
    param (
        [string]$currentValue,
        [string]$userToRemove
    )
    
    $owners = $currentValue -split ',' | ForEach-Object { $_.Trim() }
    $remainingOwners = $owners | Where-Object { $_ -ne $userToRemove }
    return ($remainingOwners -join ', ')
}

# Function to edit owner list
function Edit-OwnerList {
    param (
        [string]$currentValue
    )
    
    $owners = $currentValue -split ',' | ForEach-Object { $_.Trim() }
    Write-Host "`nCurrent owners:"
    for ($i = 0; $i -lt $owners.Count; $i++) {
        Write-Host "$($i+1): $($owners[$i])"
    }
    
    Write-Host "`nWhat would you like to do?"
    Write-Host "1: Remove an owner"
    Write-Host "2: Add a new owner"
    Write-Host "3: Replace an owner"
    $choice = Read-Host "Enter your choice (1-3)"
    
    switch ($choice) {
        "1" {
            $index = [int](Read-Host "Enter the number of the owner to remove") - 1
            if ($index -ge 0 -and $index -lt $owners.Count) {
                $owners = $owners | Where-Object { $_ -ne $owners[$index] }
            }
        }
        "2" {
            $newOwner = Read-Host "Enter the new owner's username"
            if ($owners -notcontains $newOwner) {
                $owners += $newOwner
            }
        }
        "3" {
            $index = [int](Read-Host "Enter the number of the owner to replace") - 1
            if ($index -ge 0 -and $index -lt $owners.Count) {
                $newOwner = Read-Host "Enter the new owner's username"
                $owners[$index] = $newOwner
            }
        }
    }
    
    return ($owners -join ', ')
}

# Create a log file
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

# Main script loop
do {
    $choice = Show-Menu
    switch ($choice) {
        "1" { # Remove user
            $userToRemove = Read-Host "Enter the username to remove from extensionAttribute6"
            $targetGroupNames = Get-TargetGroups
            
            foreach ($groupName in $targetGroupNames) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Properties extensionAttribute6
                    if ($group -and $group.extensionAttribute6 -like "*$userToRemove*") {
                        $currentValue = $group.extensionAttribute6
                        $newValue = Remove-FromOwnerList -currentValue $currentValue -userToRemove $userToRemove
                        
                        Write-Host "`nGroup: $groupName"
                        Write-Host "Current Value: $currentValue"
                        Write-Host "New Value: $newValue"
                        
                        $confirm = Read-Host "Update this group? (Y/N)"
                        if ($confirm -eq 'Y') {
                            if ([string]::IsNullOrWhiteSpace($newValue)) {
                                Set-ADGroup -Identity $group.DistinguishedName -Clear extensionAttribute6
                            } else {
                                Set-ADGroup -Identity $group.DistinguishedName -Replace @{extensionAttribute6 = $newValue}
                            }
                            Write-Host "Successfully updated group $groupName" -ForegroundColor Green
                            Write-ToLog "Updated group $groupName - Removed $userToRemove"
                        }
                    }
                } catch {
                    Write-Host "Error processing group $groupName: $($_.Exception.Message)" -ForegroundColor Red
                    Write-ToLog "ERROR: $groupName - $($_.Exception.Message)"
                }
            }
        }
        
        "2" { # Add user
            $userToAdd = Read-Host "Enter the username to add to extensionAttribute6"
            $targetGroupNames = Get-TargetGroups
            
            foreach ($groupName in $targetGroupNames) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Properties extensionAttribute6
                    if ($group) {
                        $currentValue = $group.extensionAttribute6
                        $newValue = Format-OwnerList -currentValue $currentValue -userToAdd $userToAdd
                        
                        Write-Host "`nGroup: $groupName"
                        Write-Host "Current Value: $currentValue"
                        Write-Host "New Value: $newValue"
                        
                        $confirm = Read-Host "Update this group? (Y/N)"
                        if ($confirm -eq 'Y') {
                            Set-ADGroup -Identity $group.DistinguishedName -Replace @{extensionAttribute6 = $newValue}
                            Write-Host "Successfully updated group $groupName" -ForegroundColor Green
                            Write-ToLog "Updated group $groupName - Added $userToAdd"
                        }
                    }
                } catch {
                    Write-Host "Error processing group $groupName: $($_.Exception.Message)" -ForegroundColor Red
                    Write-ToLog "ERROR: $groupName - $($_.Exception.Message)"
                }
            }
        }
        
        "3" { # Edit existing users
            $targetGroupNames = Get-TargetGroups
            
            foreach ($groupName in $targetGroupNames) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Properties extensionAttribute6
                    if ($group) {
                        Write-Host "`nGroup: $groupName"
                        $currentValue = $group.extensionAttribute6
                        $newValue = Edit-OwnerList -currentValue $currentValue
                        
                        Write-Host "Current Value: $currentValue"
                        Write-Host "New Value: $newValue"
                        
                        $confirm = Read-Host "Update this group? (Y/N)"
                        if ($confirm -eq 'Y') {
                            if ([string]::IsNullOrWhiteSpace($newValue)) {
                                Set-ADGroup -Identity $group.DistinguishedName -Clear extensionAttribute6
                            } else {
                                Set-ADGroup -Identity $group.DistinguishedName -Replace @{extensionAttribute6 = $newValue}
                            }
                            Write-Host "Successfully updated group $groupName" -ForegroundColor Green
                            Write-ToLog "Updated group $groupName - Edited users"
                        }
                    }
                } catch {
                    Write-Host "Error processing group $groupName: $($_.Exception.Message)" -ForegroundColor Red
                    Write-ToLog "ERROR: $groupName - $($_.Exception.Message)"
                }
            }
        }
        
        "4" { # Exit
            Write-Host "Exiting script..."
            return
        }
        
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
} while ($true)