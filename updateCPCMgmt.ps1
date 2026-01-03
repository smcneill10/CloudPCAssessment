<#
.SYNOPSIS
    Advanced Windows 365 Cloud PC Management Tool

.DESCRIPTION
    Modern PowerShell script for managing Windows 365 Cloud PCs using Microsoft Graph API.
    Supports Enterprise and Frontline (Dedicated/Shared) Cloud PC operations including:
    - Restart operations for Enterprise and Frontline Dedicated Cloud PCs
    - Power management (Start/Stop) for Frontline Dedicated Cloud PCs
    - Reprovision operations for Frontline Shared Cloud PCs
    - Bulk operations support
    - Enhanced error handling and logging

.PARAMETER CloudPCId
    The unique identifier of the Cloud PC to manage. Can be piped from Get-MgDeviceManagementVirtualEndpointCloudPc.

.PARAMETER Operation
    The operation to perform: Restart, Start, Stop, Reprovision, or List.

.PARAMETER UseDeviceCodeAuth
    Use device code flow for authentication (useful for non-interactive scenarios).

.PARAMETER TenantId
    Specify a specific tenant ID for multi-tenant scenarios.

.NOTES
    Author: Cloud PC Management Team
    Date: January 2, 2026
    Version: 2.0
    Requirements:
      - PowerShell 7.0 or later (recommended) or PowerShell 5.1
      - Microsoft.Graph.Authentication v2.0+
      - Microsoft.Graph.DeviceManagement v2.0+
    
    Permissions Required:
      - CloudPC.ReadWrite.All or
      - DeviceManagementConfiguration.ReadWrite.All

.EXAMPLE
    .\updateCPCMgmt.ps1
    Runs the interactive menu for Cloud PC management.

.EXAMPLE
    .\updateCPCMgmt.ps1 -Operation List
    Lists all Cloud PCs in the tenant.

.EXAMPLE
    .\updateCPCMgmt.ps1 -CloudPCId "abc123" -Operation Restart
    Restarts a specific Cloud PC.

.EXAMPLE
    Get-MgDeviceManagementVirtualEndpointCloudPc | Where-Object {$_.Status -eq 'Running'} | .\updateCPCMgmt.ps1 -Operation Stop
    Stops all running Cloud PCs using pipeline input.
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.DeviceManagement"; ModuleVersion="2.0.0" }

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Direct')]
    [Alias('Id')]
    [ValidatePattern('^[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}$')]
    [string]$CloudPCId,

    [Parameter(ParameterSetName = 'Direct')]
    [Parameter(ParameterSetName = 'Interactive')]
    [ValidateSet('Restart', 'Start', 'Stop', 'Reprovision', 'List', 'GetDetails')]
    [string]$Operation,

    [Parameter()]
    [switch]$UseDeviceCodeAuth,

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [ValidateSet('v1.0', 'beta')]
    [string]$GraphApiVersion = 'beta'
)

#region Configuration
$script:Config = @{
    RequiredScopes = @('CloudPC.ReadWrite.All')
    MinGraphVersion = [Version]'2.0.0'
    LogPath = Join-Path $env:TEMP "CloudPCManagement_$(Get-Date -Format 'yyyyMMdd').log"
    Colors = @{
        Success = 'Green'
        Warning = 'Yellow'
        Error = 'Red'
        Info = 'Cyan'
        Header = 'Magenta'
    }
}
#endregion

#region Helper Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes messages to console and log file with color coding.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    try {
        Add-Content -Path $script:Config.LogPath -Value $logMessage -ErrorAction SilentlyContinue
    } catch {
        # Silently fail if can't write to log
    }
    
    # Write to console with color
    $color = $script:Config.Colors[$Level]
    switch ($Level) {
        'Error'   { Write-Error $Message }
        'Warning' { Write-Warning $Message }
        default   { Write-Host $Message -ForegroundColor $color }
    }
}

function Test-GraphConnection {
    <#
    .SYNOPSIS
        Verifies Microsoft Graph connection and required permissions.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    try {
        $context = Get-MgContext -ErrorAction Stop
        
        if (-not $context) {
            Write-Log "Not connected to Microsoft Graph" -Level Warning
            return $false
        }
        
        # Verify required scopes
        $hasRequiredScope = $script:Config.RequiredScopes | ForEach-Object {
            $context.Scopes -contains $_
        }
        
        if ($hasRequiredScope -notcontains $true) {
            Write-Log "Missing required scope: $($script:Config.RequiredScopes -join ', ')" -Level Warning
            return $false
        }
        
        Write-Log "Connected to tenant: $($context.TenantId) as $($context.Account)" -Level Success
        return $true
        
    } catch {
        Write-Log "Error checking Graph connection: $_" -Level Error
        return $false
    }
}

function Initialize-GraphConnection {
    <#
    .SYNOPSIS
        Establishes connection to Microsoft Graph with proper scopes.
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Initializing Microsoft Graph connection..." -Level Info
    
    $connectParams = @{
        Scopes = $script:Config.RequiredScopes
        NoWelcome = $true
        ErrorAction = 'Stop'
    }
    
    if ($UseDeviceCodeAuth) {
        $connectParams['UseDeviceCode'] = $true
        Write-Log "Using Device Code authentication flow" -Level Info
    }
    
    if ($TenantId) {
        $connectParams['TenantId'] = $TenantId
    }
    
    try {
        Connect-MgGraph @connectParams
        
        $context = Get-MgContext
        Write-Log "Successfully connected to Microsoft Graph" -Level Success
        Write-Log "  Tenant: $($context.TenantId)" -Level Info
        Write-Log "  Account: $($context.Account)" -Level Info
        Write-Log "  Scopes: $($context.Scopes -join ', ')" -Level Info
        
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level Error
        throw
    }
}

function Disconnect-GraphConnection {
    <#
    .SYNOPSIS
        Safely disconnects from Microsoft Graph.
    #>
    [CmdletBinding()]
    param()
    
    try {
        if (Get-MgContext) {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Write-Log "Disconnected from Microsoft Graph" -Level Info
        }
    } catch {
        Write-Log "Error disconnecting from Graph: $_" -Level Warning
    }
}

#endregion

#region Cloud PC Management Functions

function Get-CloudPCDetails {
    <#
    .SYNOPSIS
        Retrieves detailed information about a Cloud PC.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CloudPCId
    )
    
    try {
        Write-Log "Retrieving details for Cloud PC: $CloudPCId" -Level Info
        
        $cloudPC = Get-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
        
        if ($cloudPC) {
            Write-Log "Cloud PC Details:" -Level Success
            Write-Log "  Display Name: $($cloudPC.DisplayName)" -Level Info
            Write-Log "  Status: $($cloudPC.Status)" -Level Info
            Write-Log "  Type: $($cloudPC.ProvisioningType)" -Level Info
            Write-Log "  User: $($cloudPC.UserPrincipalName)" -Level Info
            Write-Log "  Last Modified: $($cloudPC.LastModifiedDateTime)" -Level Info
            Write-Log "  Grace Period End: $($cloudPC.GracePeriodEndDateTime)" -Level Info
            
            return $cloudPC
        }
        
    } catch {
        Write-Log "Failed to retrieve Cloud PC details: $_" -Level Error
        throw
    }
}

function Get-AllCloudPCs {
    <#
    .SYNOPSIS
        Lists all Cloud PCs in the tenant.
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Retrieving all Cloud PCs..." -Level Info
        
        $cloudPCs = Get-MgDeviceManagementVirtualEndpointCloudPc -All -ErrorAction Stop
        
        if ($cloudPCs) {
            Write-Log "Found $($cloudPCs.Count) Cloud PC(s)" -Level Success
            
            $cloudPCs | Format-Table -Property `
                @{Label='Display Name'; Expression={$_.DisplayName}},
                @{Label='Status'; Expression={$_.Status}},
                @{Label='Type'; Expression={$_.ProvisioningType}},
                @{Label='User'; Expression={$_.UserPrincipalName}},
                @{Label='Cloud PC ID'; Expression={$_.Id}} -AutoSize
            
            return $cloudPCs
        } else {
            Write-Log "No Cloud PCs found in the tenant" -Level Warning
        }
        
    } catch {
        Write-Log "Failed to retrieve Cloud PCs: $_" -Level Error
        throw
    }
}

function Restart-CloudPC {
    <#
    .SYNOPSIS
        Restarts a Cloud PC (Enterprise or Frontline Dedicated).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$CloudPCId
    )
    
    try {
        # Get Cloud PC details first
        $cloudPC = Get-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
        
        if ($PSCmdlet.ShouldProcess($cloudPC.DisplayName, "Restart Cloud PC")) {
            Write-Log "Restarting Cloud PC: $($cloudPC.DisplayName) ($CloudPCId)" -Level Info
            
            Restart-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
            
            Write-Log "Restart command issued successfully" -Level Success
            return $true
        }
        
    } catch {
        Write-Log "Failed to restart Cloud PC: $_" -Level Error
        return $false
    }
}

function Start-CloudPC {
    <#
    .SYNOPSIS
        Starts a Frontline Dedicated Cloud PC.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$CloudPCId
    )
    
    try {
        $cloudPC = Get-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
        
        if ($cloudPC.ProvisioningType -ne 'Dedicated') {
            Write-Log "Start operation is only supported for Frontline Dedicated Cloud PCs" -Level Warning
            return $false
        }
        
        if ($PSCmdlet.ShouldProcess($cloudPC.DisplayName, "Start Cloud PC")) {
            Write-Log "Starting Cloud PC: $($cloudPC.DisplayName) ($CloudPCId)" -Level Info
            
            Start-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
            
            Write-Log "Start command issued successfully" -Level Success
            return $true
        }
        
    } catch {
        Write-Log "Failed to start Cloud PC: $_" -Level Error
        return $false
    }
}

function Stop-CloudPC {
    <#
    .SYNOPSIS
        Stops a Frontline Dedicated Cloud PC.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$CloudPCId
    )
    
    try {
        $cloudPC = Get-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
        
        if ($cloudPC.ProvisioningType -ne 'Dedicated') {
            Write-Log "Stop operation is only supported for Frontline Dedicated Cloud PCs" -Level Warning
            return $false
        }
        
        if ($PSCmdlet.ShouldProcess($cloudPC.DisplayName, "Stop Cloud PC")) {
            Write-Log "Stopping Cloud PC: $($cloudPC.DisplayName) ($CloudPCId)" -Level Info
            
            Stop-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
            
            Write-Log "Stop command issued successfully" -Level Success
            return $true
        }
        
    } catch {
        Write-Log "Failed to stop Cloud PC: $_" -Level Error
        return $false
    }
}

function Invoke-CloudPCReprovision {
    <#
    .SYNOPSIS
        Reprovisions a Frontline Shared Cloud PC.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$CloudPCId
    )
    
    try {
        $cloudPC = Get-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
        
        if ($PSCmdlet.ShouldProcess($cloudPC.DisplayName, "Reprovision Cloud PC")) {
            Write-Log "Reprovisioning Cloud PC: $($cloudPC.DisplayName) ($CloudPCId)" -Level Info
            Write-Log "WARNING: This will reset the Cloud PC to its original state" -Level Warning
            
            Invoke-MgReprovisionDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId -ErrorAction Stop
            
            Write-Log "Reprovision command issued successfully" -Level Success
            return $true
        }
        
    } catch {
        Write-Log "Failed to reprovision Cloud PC: $_" -Level Error
        return $false
    }
}

#endregion

#region Interactive Menu

function Show-InteractiveMenu {
    <#
    .SYNOPSIS
        Displays interactive menu for Cloud PC management.
    #>
    [CmdletBinding()]
    param()
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor $script:Config.Colors.Header
        Write-Host "  Windows 365 Cloud PC Management Tool  " -ForegroundColor $script:Config.Colors.Header
        Write-Host "           Version 2.0 (2026)           " -ForegroundColor $script:Config.Colors.Header
        Write-Host "========================================" -ForegroundColor $script:Config.Colors.Header
        Write-Host ""
        Write-Host "1. List all Cloud PCs" -ForegroundColor $script:Config.Colors.Info
        Write-Host "2. Get Cloud PC details" -ForegroundColor $script:Config.Colors.Info
        Write-Host "3. Restart Cloud PC" -ForegroundColor $script:Config.Colors.Info
        Write-Host "4. Start Frontline Dedicated Cloud PC" -ForegroundColor $script:Config.Colors.Info
        Write-Host "5. Stop Frontline Dedicated Cloud PC" -ForegroundColor $script:Config.Colors.Info
        Write-Host "6. Reprovision Frontline Shared Cloud PC" -ForegroundColor $script:Config.Colors.Info
        Write-Host "7. Reconnect to Microsoft Graph" -ForegroundColor $script:Config.Colors.Info
        Write-Host "8. Exit" -ForegroundColor $script:Config.Colors.Info
        Write-Host ""
        
        $choice = Read-Host "Enter your selection (1-8)"
        
        switch ($choice) {
            "1" {
                Get-AllCloudPCs
                Pause
            }
            "2" {
                $id = Read-Host "Enter the Cloud PC ID"
                if ($id) {
                    Get-CloudPCDetails -CloudPCId $id
                }
                Pause
            }
            "3" {
                $id = Read-Host "Enter the Cloud PC ID to restart"
                if ($id) {
                    Restart-CloudPC -CloudPCId $id -Confirm
                }
                Pause
            }
            "4" {
                $id = Read-Host "Enter the Frontline Dedicated Cloud PC ID to start"
                if ($id) {
                    Start-CloudPC -CloudPCId $id -Confirm
                }
                Pause
            }
            "5" {
                $id = Read-Host "Enter the Frontline Dedicated Cloud PC ID to stop"
                if ($id) {
                    Stop-CloudPC -CloudPCId $id -Confirm
                }
                Pause
            }
            "6" {
                $id = Read-Host "Enter the Frontline Shared Cloud PC ID to reprovision"
                if ($id) {
                    Write-Host "WARNING: Reprovision will reset the Cloud PC to its original state!" -ForegroundColor Yellow
                    $confirm = Read-Host "Are you sure you want to continue? (yes/no)"
                    if ($confirm -eq 'yes') {
                        Invoke-CloudPCReprovision -CloudPCId $id
                    }
                }
                Pause
            }
            "7" {
                Disconnect-GraphConnection
                Initialize-GraphConnection
                Pause
            }
            "8" {
                Write-Log "Exiting Cloud PC Management Tool" -Level Info
                break
            }
            default {
                Write-Log "Invalid selection. Please choose 1-8." -Level Warning
                Start-Sleep -Seconds 2
            }
        }
        
    } while ($choice -ne "8")
}

function Pause {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Main Execution

try {
    Write-Log "Starting Windows 365 Cloud PC Management Tool" -Level Info
    
    # Verify or establish Graph connection
    if (-not (Test-GraphConnection)) {
        Initialize-GraphConnection
    }
    
    # Handle different execution modes
    if ($PSCmdlet.ParameterSetName -eq 'Direct' -and $Operation) {
        # Direct operation mode
        switch ($Operation) {
            'List' {
                Get-AllCloudPCs
            }
            'GetDetails' {
                if (-not $CloudPCId) {
                    throw "CloudPCId is required for GetDetails operation"
                }
                Get-CloudPCDetails -CloudPCId $CloudPCId
            }
            'Restart' {
                if (-not $CloudPCId) {
                    throw "CloudPCId is required for Restart operation"
                }
                Restart-CloudPC -CloudPCId $CloudPCId
            }
            'Start' {
                if (-not $CloudPCId) {
                    throw "CloudPCId is required for Start operation"
                }
                Start-CloudPC -CloudPCId $CloudPCId
            }
            'Stop' {
                if (-not $CloudPCId) {
                    throw "CloudPCId is required for Stop operation"
                }
                Stop-CloudPC -CloudPCId $CloudPCId
            }
            'Reprovision' {
                if (-not $CloudPCId) {
                    throw "CloudPCId is required for Reprovision operation"
                }
                Invoke-CloudPCReprovision -CloudPCId $CloudPCId
            }
        }
    } else {
        # Interactive mode
        Show-InteractiveMenu
    }
    
} catch {
    Write-Log "Critical error: $_" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
    
} finally {
    # Cleanup
    Write-Log "Log file saved to: $($script:Config.LogPath)" -Level Info
    
    # Optionally disconnect (comment out if you want to keep the session)
    # Disconnect-GraphConnection
}

#endregion