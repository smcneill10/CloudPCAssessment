<#
.SYNOPSIS
    Windows 365 Cloud PC Management Tool

.DESCRIPTION
    Interactive PowerShell tool for managing Windows 365 Cloud PCs using Microsoft Graph API.
    Supports viewing, starting, stopping, and restarting Cloud PCs with modern best practices.

.NOTES
    Version:        2.0
    Author:         Updated for modern best practices
    Last Updated:   January 2026
    
    Requirements:
    - PowerShell 7.0 or later
    - Microsoft.Graph.Beta.DeviceManagement.Administration module
    - Appropriate Microsoft Graph permissions

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/cloudpc
    
.EXAMPLE
    .\updateCPCAssessment.ps1
    Runs the interactive Cloud PC management tool
#>

#Requires -Version 7.0

[CmdletBinding()]
param()

#region Module Check
# Verify required modules are installed (PowerShell will auto-load them when needed)
$requiredModules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Beta.DeviceManagement.Administration'
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        Write-Warning "Required module '$module' is not installed."
        Write-Host "Install it with: Install-Module -Name $module -Scope CurrentUser -Force" -ForegroundColor Yellow
        exit 1
    }
}

Write-Verbose "All required modules are available. PowerShell will auto-load them as needed."
#endregion

#region Configuration
$script:Config = @{
    RequiredScopes = @(
        'CloudPC.ReadWrite.All'
        'CloudPC.Read.All'
        'User.Read.All'
        'Group.Read.All'
    )
    Colors = @{
        Success = 'Green'
        Error   = 'Red'
        Info    = 'Cyan'
        Warning = 'Yellow'
        Normal  = 'White'
    }
}
#endregion

#region Helper Functions

function Write-ColorMessage {
    <#
    .SYNOPSIS
        Writes colored messages to the console
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Success', 'Error', 'Info', 'Warning', 'Normal')]
        [string]$Type = 'Normal'
    )
    
    $color = $script:Config.Colors[$Type]
    Write-Host $Message -ForegroundColor $color
}

function Initialize-GraphConnection {
    <#
    .SYNOPSIS
        Initializes and validates Microsoft Graph connection
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-ColorMessage "Connecting to Microsoft Graph..." -Type Info
        
        # Connect with required scopes
        Connect-MgGraph -Scopes $script:Config.RequiredScopes -NoWelcome -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if ($null -eq $context) {
            throw "Failed to establish Graph connection"
        }
        
        Write-ColorMessage "Successfully connected to Microsoft Graph" -Type Success
        Write-ColorMessage "Account: $($context.Account)" -Type Info
        Write-ColorMessage "Tenant: $($context.TenantId)" -Type Info
        Write-Host ""
        
        return $true
    }
    catch {
        Write-ColorMessage "Error connecting to Microsoft Graph: $_" -Type Error
        return $false
    }
}

function Get-CloudPCList {
    <#
    .SYNOPSIS
        Retrieves all Cloud PCs with comprehensive properties
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-ColorMessage "Retrieving Cloud PCs..." -Type Info
        
        # Using Beta API for full feature set (PowerState, Status, etc.)
        $cloudPCs = Get-MgBetaDeviceManagementVirtualEndpointCloudPc -All -ErrorAction Stop
        
        if ($null -eq $cloudPCs -or $cloudPCs.Count -eq 0) {
            Write-ColorMessage "No Cloud PCs found in this tenant" -Type Warning
            return $null
        }
        
        Write-ColorMessage "Found $($cloudPCs.Count) Cloud PC(s)" -Type Success
        return $cloudPCs
    }
    catch {
        Write-ColorMessage "Error retrieving Cloud PCs: $_" -Type Error
        return $null
    }
}

function Show-CloudPCMenu {
    <#
    .SYNOPSIS
        Displays the Cloud PC selection menu
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$CloudPCs
    )
    
    Write-Host ""
    Write-ColorMessage "=== Available Cloud PCs ===" -Type Info
    Write-Host ""
    
    for ($i = 0; $i -lt $CloudPCs.Count; $i++) {
        $cpc = $CloudPCs[$i]
        $index = $i + 1
        $displayName = $cpc.ManagedDeviceName ?? $cpc.DisplayName
        $user = $cpc.UserPrincipalName
        $powerState = $cpc.PowerState ?? "Unknown"
        $status = $cpc.Status ?? "Unknown"
        
        # Color code based on power state
        $stateColor = switch ($powerState) {
            'running' { 'Success' }
            'stopped' { 'Warning' }
            'poweredOff' { 'Warning' }
            default { 'Normal' }
        }
        
        Write-Host "  [$index] " -NoNewline
        Write-Host "$displayName " -ForegroundColor White -NoNewline
        Write-Host "| User: $user " -ForegroundColor Gray -NoNewline
        Write-Host "| Status: $status " -ForegroundColor Cyan -NoNewline
        Write-ColorMessage "| Power: $powerState" -Type $stateColor
    }
    
    Write-Host ""
    Write-ColorMessage "  [0] Exit" -Type Warning
    Write-Host ""
}

function Show-CloudPCDetails {
    <#
    .SYNOPSIS
        Displays detailed information about a specific Cloud PC
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$CloudPC
    )
    
    Write-Host ""
    Write-ColorMessage "=== Cloud PC Details ===" -Type Info
    Write-Host ""
    
    $details = [ordered]@{
        'Display Name'           = $CloudPC.DisplayName
        'User Principal Name'    = $CloudPC.UserPrincipalName
        'Managed Device Name'    = $CloudPC.ManagedDeviceName
        'Cloud PC ID'            = $CloudPC.Id
        'Status'                 = $CloudPC.Status
        'Power State'            = $CloudPC.PowerState
        'Service Plan'           = $CloudPC.ServicePlanName
        'Provisioning Policy'    = $CloudPC.ProvisioningPolicyName
        'Image Name'             = $CloudPC.ImageDisplayName
        'Last Modified'          = $CloudPC.LastModifiedDateTime
        'Grace Period End'       = $CloudPC.GracePeriodEndDateTime
    }
    
    foreach ($key in $details.Keys) {
        $value = $details[$key] ?? 'N/A'
        Write-Host "  " -NoNewline
        Write-Host "$key" -ForegroundColor Cyan -NoNewline
        Write-Host ": " -NoNewline
        Write-Host "$value" -ForegroundColor White
    }
    
    Write-Host ""
}

function Show-ActionMenu {
    <#
    .SYNOPSIS
        Displays the action menu for a selected Cloud PC
    #>
    [CmdletBinding()]
    param()
    
    Write-ColorMessage "=== Available Actions ===" -Type Info
    Write-Host ""
    Write-Host "  [1] Start Cloud PC"
    Write-Host "  [2] Stop Cloud PC"
    Write-Host "  [3] Restart Cloud PC"
    Write-Host "  [4] Refresh Details"
    Write-Host "  [5] Back to Cloud PC List"
    Write-Host "  [0] Exit"
    Write-Host ""
}

function Invoke-CloudPCAction {
    <#
    .SYNOPSIS
        Executes the selected action on a Cloud PC
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Action,
        
        [Parameter(Mandatory)]
        [object]$CloudPC
    )
    
    try {
        switch ($Action) {
            '1' {
                Write-ColorMessage "Starting Cloud PC: $($CloudPC.DisplayName)..." -Type Info
                Start-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPC.Id -ErrorAction Stop
                Write-ColorMessage "Start command sent successfully" -Type Success
            }
            '2' {
                Write-ColorMessage "Stopping Cloud PC: $($CloudPC.DisplayName)..." -Type Warning
                Stop-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPC.Id -ErrorAction Stop
                Write-ColorMessage "Stop command sent successfully" -Type Success
            }
            '3' {
                Write-ColorMessage "Restarting Cloud PC: $($CloudPC.DisplayName)..." -Type Warning
                Restart-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPC.Id -ErrorAction Stop
                Write-ColorMessage "Restart command sent successfully" -Type Success
            }
            '4' {
                # Refresh details
                $refreshed = Get-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPC.Id -ErrorAction Stop
                Show-CloudPCDetails -CloudPC $refreshed
                return $refreshed
            }
        }
        
        Start-Sleep -Seconds 2
        return $CloudPC
    }
    catch {
        Write-ColorMessage "Error executing action: $_" -Type Error
        Start-Sleep -Seconds 2
        return $CloudPC
    }
}

function Start-CloudPCManagement {
    <#
    .SYNOPSIS
        Main management loop for Cloud PC operations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$CloudPCs
    )
    
    do {
        Clear-Host
        Show-CloudPCMenu -CloudPCs $CloudPCs
        
        $selection = Read-Host "Select a Cloud PC (0 to exit)"
        
        if ($selection -eq '0') {
            Write-ColorMessage "Exiting..." -Type Info
            break
        }
        
        if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $CloudPCs.Count) {
            $selectedIndex = [int]$selection - 1
            $selectedCloudPC = $CloudPCs[$selectedIndex]
            
            # Inner loop for selected Cloud PC
            do {
                Clear-Host
                Show-CloudPCDetails -CloudPC $selectedCloudPC
                Show-ActionMenu
                
                $actionSelection = Read-Host "Select an action"
                
                if ($actionSelection -eq '0') {
                    Write-ColorMessage "Exiting..." -Type Info
                    return
                }
                elseif ($actionSelection -eq '5') {
                    break  # Back to Cloud PC list
                }
                elseif ($actionSelection -in @('1', '2', '3', '4')) {
                    $selectedCloudPC = Invoke-CloudPCAction -Action $actionSelection -CloudPC $selectedCloudPC
                }
                else {
                    Write-ColorMessage "Invalid selection. Please try again." -Type Warning
                    Start-Sleep -Seconds 1
                }
                
            } while ($true)
        }
        else {
            Write-ColorMessage "Invalid selection. Please try again." -Type Warning
            Start-Sleep -Seconds 1
        }
        
    } while ($true)
}

#endregion

#region Main Execution

function Start-Main {
    <#
    .SYNOPSIS
        Main entry point for the script
    #>
    [CmdletBinding()]
    param()
    
    try {
        Clear-Host
        Write-ColorMessage "=== Windows 365 Cloud PC Management Tool ===" -Type Info
        Write-Host ""
        
        # Initialize Graph connection
        if (-not (Initialize-GraphConnection)) {
            Write-ColorMessage "Unable to proceed without Graph connection" -Type Error
            return
        }
        
        # Get Cloud PCs
        $cloudPCs = Get-CloudPCList
        if ($null -eq $cloudPCs) {
            Write-ColorMessage "No Cloud PCs available to manage" -Type Warning
            return
        }
        
        # Start management interface
        Start-CloudPCManagement -CloudPCs $cloudPCs
        
        Write-ColorMessage "Thank you for using the Cloud PC Management Tool!" -Type Success
    }
    catch {
        Write-ColorMessage "An unexpected error occurred: $_" -Type Error
    }
    finally {
        # Cleanup if needed
        Write-Verbose "Script execution completed"
    }
}

# Execute main function
Start-Main

#endregion
