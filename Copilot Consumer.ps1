<# 
.SYNOPSIS
    Manage Windows 365 Cloud PCs for Enterprise and Frontline scenarios.

.DESCRIPTION
    This PowerShell script provides management functionalities for Windows 365 Cloud PCs:
      • For Enterprise Cloud PCs: Issues a reboot command.
      • For Frontline Dedicated Cloud PCs: Supports starting, stopping, and rebooting.
      • For Frontline Shared Cloud PCs: Offers a reprovision operation.
      
    The script leverages the Microsoft Graph API; it requires that you are connected using
    valid credentials with the proper permissions (such as DeviceManagementManagedDevices.ReadWrite.All).

.NOTES
    Author: [Your Name]
    Date: [Current Date]
    Requirements:
      - PowerShell 5.1 or later.
      - Microsoft.Graph module installed.
         Install via: Install-Module Microsoft.Graph -Scope CurrentUser
    Usage:
      Run the script and follow the on-screen prompts to select which action you want to perform.
      
.EXAMPLE
    PS C:\> .\Manage-Windows365CloudPCs.ps1
    (Follow the prompts to manage your Cloud PCs.)
#>

# -------------------------------------------------------------------------------------------------
# Import the Microsoft Graph module.
# If not installed, instruct the user to install it.
try {
    Import-Module Microsoft.Graph -ErrorAction Stop
} catch {
    Write-Error "The Microsoft.Graph module is not installed. Please run `Install-Module Microsoft.Graph -Scope CurrentUser` and try again."
    exit
}

# -------------------------------------------------------------------------------------------------
# Function: Initialize-Windows365Session
# This function connects to Microsoft Graph so that subsequent Cloud PC management commands are authenticated.
function Initialize-Windows365Session {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph for a Windows 365 management session.
        
    .DESCRIPTION
        Uses the Connect-MgGraph cmdlet with the required permission scope to ensure that
        operations such as reboot, start, stop, and reprovision can be successfully issued.
        
    .EXAMPLE
        Initialize-Windows365Session
    #>
    Write-Output "Connecting to Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All" -ErrorAction Stop
        Write-Output "Successfully connected to Microsoft Graph."
    } catch {
        Write-Error "Failed to connect to Microsoft Graph. Please verify your credentials and permission scopes."
        exit
    }
}

# -------------------------------------------------------------------------------------------------
# Function: Reboot-EnterpriseCPC
# This function reboots an Enterprise Cloud PC identified by its unique ID.
function Reboot-EnterpriseCPC {
    <#
    .SYNOPSIS
        Reboots an Enterprise Cloud PC.
        
    .DESCRIPTION
        Sends a reboot command to the specified Enterprise Cloud PC via the Microsoft Graph API.
        Replace <CloudPCId> with the appropriate identifier of the target Cloud PC.
        
    .PARAMETER CloudPCId
        The unique identifier of the Enterprise Cloud PC to be rebooted.
        
    .EXAMPLE
        Reboot-EnterpriseCPC -CloudPCId "12345-abcd"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId
    )

    # Construct the endpoint URI.
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windows365CloudPCs/$CloudPCId/reboot"
    Write-Output "Attempting to reboot Enterprise Cloud PC with ID: $CloudPCId"
    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        Write-Output "Reboot command issued successfully for Cloud PC ID: $CloudPCId"
    } catch {
        Write-Error "Failed to reboot Cloud PC ID: $CloudPCId. Error details: $_"
    }
}

# -------------------------------------------------------------------------------------------------
# Function: Start-FrontlineDedicatedCPC
# This function starts a Frontline Dedicated Cloud PC using its unique ID.
function Start-FrontlineDedicatedCPC {
    <#
    .SYNOPSIS
        Starts a Frontline Dedicated Cloud PC.
        
    .DESCRIPTION
        Uses the Microsoft Graph API to send a start command to the specified Frontline Dedicated Cloud PC.
        
    .PARAMETER CloudPCId
        The unique identifier of the Frontline Dedicated Cloud PC to start.
        
    .EXAMPLE
        Start-FrontlineDedicatedCPC -CloudPCId "67890-efgh"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId
    )
    
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windows365CloudPCs/$CloudPCId/start"
    Write-Output "Attempting to start Frontline Dedicated Cloud PC with ID: $CloudPCId"
    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        Write-Output "Start command issued successfully for Cloud PC ID: $CloudPCId"
    } catch {
        Write-Error "Failed to start Cloud PC ID: $CloudPCId. Error details: $_"
    }
}

# -------------------------------------------------------------------------------------------------
# Function: Stop-FrontlineDedicatedCPC
# This function stops a Frontline Dedicated Cloud PC as identified by its unique ID.
function Stop-FrontlineDedicatedCPC {
    <#
    .SYNOPSIS
        Stops a Frontline Dedicated Cloud PC.
        
    .DESCRIPTION
        Sends a stop command via the Microsoft Graph API to the specified Frontline Dedicated Cloud PC.
        
    .PARAMETER CloudPCId
        The unique identifier of the Frontline Dedicated Cloud PC to stop.
        
    .EXAMPLE
        Stop-FrontlineDedicatedCPC -CloudPCId "67890-efgh"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId
    )
    
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windows365CloudPCs/$CloudPCId/stop"
    Write-Output "Attempting to stop Frontline Dedicated Cloud PC with ID: $CloudPCId"
    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        Write-Output "Stop command issued successfully for Cloud PC ID: $CloudPCId"
    } catch {
        Write-Error "Failed to stop Cloud PC ID: $CloudPCId. Error details: $_"
    }
}

# -------------------------------------------------------------------------------------------------
# Function: Reboot-FrontlineDedicatedCPC
# This function reboots a Frontline Dedicated Cloud PC via its unique ID.
function Reboot-FrontlineDedicatedCPC {
    <#
    .SYNOPSIS
        Reboots a Frontline Dedicated Cloud PC.
        
    .DESCRIPTION
        Issues a reboot command using the Microsoft Graph API to the specified Frontline Dedicated Cloud PC.
        
    .PARAMETER CloudPCId
        The unique identifier of the Frontline Dedicated Cloud PC to reboot.
        
    .EXAMPLE
        Reboot-FrontlineDedicatedCPC -CloudPCId "67890-efgh"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId
    )
    
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windows365CloudPCs/$CloudPCId/reboot"
    Write-Output "Attempting to reboot Frontline Dedicated Cloud PC with ID: $CloudPCId"
    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        Write-Output "Reboot command issued successfully for Cloud PC ID: $CloudPCId"
    } catch {
        Write-Error "Failed to reboot Cloud PC ID: $CloudPCId. Error details: $_"
    }
}

# -------------------------------------------------------------------------------------------------
# Function: Reprovision-FrontlineSharedCPC
# This function reprovisions a Frontline Shared Cloud PC identified by its unique ID.
function Reprovision-FrontlineSharedCPC {
    <#
    .SYNOPSIS
        Reprovisions a Frontline Shared Cloud PC.
        
    .DESCRIPTION
        Initiates the reprovision process using the Microsoft Graph API to essentially reset the
        specified Frontline Shared Cloud PC.
        
    .PARAMETER CloudPCId
        The unique identifier of the Frontline Shared Cloud PC to reprovision.
        
    .EXAMPLE
        Reprovision-FrontlineSharedCPC -CloudPCId "abcdef-12345"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId
    )
    
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windows365CloudPCs/$CloudPCId/reprovision"
    Write-Output "Attempting to reprovision Frontline Shared Cloud PC with ID: $CloudPCId"
    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        Write-Output "Reprovision command issued successfully for Cloud PC ID: $CloudPCId"
    } catch {
        Write-Error "Failed to reprovision Cloud PC ID: $CloudPCId. Error details: $_"
    }
}

# -------------------------------------------------------------------------------------------------
# Interactive Menu to select and perform actions.
function Show-Menu {
    Write-Host "======================================"
    Write-Host "Windows 365 Cloud PC Management Script"
    Write-Host "======================================"
    Write-Host "1. Reboot Enterprise Cloud PC"
    Write-Host "2. Start Frontline Dedicated Cloud PC"
    Write-Host "3. Stop Frontline Dedicated Cloud PC"
    Write-Host "4. Reboot Frontline Dedicated Cloud PC"
    Write-Host "5. Reprovision Frontline Shared Cloud PC"
    Write-Host "6. Exit"
}

# -------------------------------------------------------------------------------------------------
# Main script execution.
try {
    # Connect to Microsoft Graph.
    Initialize-Windows365Session

    do {
        Show-Menu
        $choice = Read-Host "Enter your selection (1-6)"
        switch ($choice) {
            "1" {
                $cloudPCId = Read-Host "Enter the Enterprise Cloud PC ID to reboot"
                Reboot-EnterpriseCPC -CloudPCId $cloudPCId
            }
            "2" {
                $cloudPCId = Read-Host "Enter the Frontline Dedicated Cloud PC ID to start"
                Start-FrontlineDedicatedCPC -CloudPCId $cloudPCId
            }
            "3" {
                $cloudPCId = Read-Host "Enter the Frontline Dedicated Cloud PC ID to stop"
                Stop-FrontlineDedicatedCPC -CloudPCId $cloudPCId
            }
            "4" {
                $cloudPCId = Read-Host "Enter the Frontline Dedicated Cloud PC ID to reboot"
                Reboot-FrontlineDedicatedCPC -CloudPCId $cloudPCId
            }
            "5" {
                $cloudPCId = Read-Host "Enter the Frontline Shared Cloud PC ID to reprovision"
                Reprovision-FrontlineSharedCPC -CloudPCId $cloudPCId
            }
            "6" {
                Write-Host "Exiting the script. Goodbye!"
                break
            }
            default {
                Write-Host "Invalid selection. Please choose a valid option (1-6)."
            }
        }
        # Pause before showing the menu again, unless exiting.
        if ($choice -ne "6") {
            Write-Host "`nPress Enter to continue..."
            [void][System.Console]::ReadLine()
        }
    } while ($choice -ne "6")
} catch {
    Write-Error "An unexpected error occurred: $_"
}