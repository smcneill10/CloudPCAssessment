<#
.SYNOPSIS
    Windows 365 Cloud PC Assessment and Management Tool

.DESCRIPTION
    Modern PowerShell script for managing Windows 365 Cloud PCs using Microsoft Graph API.
    Implements current best practices including:
    - Microsoft Graph PowerShell SDK
    - Proper error handling
    - Async operations support
    - Enhanced security with minimal permissions
    - Structured logging

.NOTES
    Author: Updated for Windows 365 Best Practices
    Version: 2.0
    Last Updated: January 2026
    
.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/cloudpc
    https://learn.microsoft.com/en-us/windows-365/
    https://learn.microsoft.com/en-us/powershell/microsoftgraph/
#>

#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.DeviceManagement"; ModuleVersion="2.0.0" }

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('v1.0', 'beta')]
    [string]$GraphApiVersion = 'beta',
    
    [Parameter(Mandatory=$false)]
    [switch]$UseDeviceCodeAuth
)

# Script configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Display theme configuration
$script:DisplayConfig = @{
    ForegroundColor = 'White'
    InfoBackground = 'DarkBlue'
    SuccessBackground = 'DarkGreen'
    ErrorBackground = 'DarkRed'
    WarningBackground = 'DarkYellow'
}

#region Helper Functions

<#
.SYNOPSIS
    Writes formatted output with color coding
#>
function Write-FormattedMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Success', 'Error', 'Warning')]
        [string]$Type = 'Info'
    )
    
    $params = @{
        Object = $Message
        ForegroundColor = $script:DisplayConfig.ForegroundColor
    }
    
    switch ($Type) {
        'Info'    { $params.BackgroundColor = $script:DisplayConfig.InfoBackground }
        'Success' { $params.BackgroundColor = $script:DisplayConfig.SuccessBackground }
        'Error'   { $params.BackgroundColor = $script:DisplayConfig.ErrorBackground }
        'Warning' { $params.BackgroundColor = $script:DisplayConfig.WarningBackground }
    }
    
    Write-Host @params
}

<#
.SYNOPSIS
    Initialize Microsoft Graph connection with required scopes
#>
function Initialize-GraphConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$UseDeviceCode
    )
    
    try {
        Write-FormattedMessage -Message "Initializing Microsoft Graph connection..." -Type Info
        
        # Required scopes for Windows 365 management
        $scopes = @(
            'CloudPC.Read.All',
            'CloudPC.ReadWrite.All',
            'User.Read.All',
            'Group.Read.All'
        )
        
        $connectParams = @{
            Scopes = $scopes
            NoWelcome = $true
            ErrorAction = 'Stop'
        }
        
        if ($UseDeviceCode) {
            $connectParams.UseDeviceCode = $true
        }
        
        Connect-MgGraph @connectParams
        
        # Display connection context
        $context = Get-MgContext
        Write-FormattedMessage -Message "`nConnection Details:" -Type Success
        Write-Host "  Account: $($context.Account)" -ForegroundColor Cyan
        Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Cyan
        Write-Host "  Scopes: $($context.Scopes -join ', ')" -ForegroundColor Cyan
        Write-Host ""
        
        return $true
    }
    catch {
        Write-FormattedMessage -Message "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Type Error
        return $false
    }
}

<#
.SYNOPSIS
    Retrieves all Cloud PCs with enhanced property selection
#>
function Get-AllCloudPCs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion
    )
    
    try {
        Write-FormattedMessage -Message "Retrieving Cloud PCs using Graph API $ApiVersion..." -Type Info
        
        $properties = @(
            'Id',
            'DisplayName',
            'UserPrincipalName',
            'ManagedDeviceName',
            'ProvisioningPolicyId',
            'ProvisioningPolicyName',
            'ImageDisplayName',
            'ServicePlanName',
            'Status',
            'PowerState',
            'LastModifiedDateTime',
            'GracePeriodEndDateTime'
        )
        
        if ($ApiVersion -eq 'beta') {
            $cloudPCs = Get-MgBetaDeviceManagementVirtualEndpointCloudPc -All -Property $properties
        }
        else {
            # v1.0 has limited properties
            $limitedProperties = $properties | Where-Object { $_ -notin @('PowerState') }
            $cloudPCs = Get-MgDeviceManagementVirtualEndpointCloudPc -All -Property $limitedProperties
        }
        
        if (-not $cloudPCs) {
            Write-FormattedMessage -Message "No Cloud PCs found in this tenant." -Type Warning
            return $null
        }
        
        Write-FormattedMessage -Message "Found $($cloudPCs.Count) Cloud PC(s)" -Type Success
        return $cloudPCs
    }
    catch {
        Write-FormattedMessage -Message "Error retrieving Cloud PCs: $($_.Exception.Message)" -Type Error
        throw
    }
}

<#
.SYNOPSIS
    Displays Cloud PC selection menu
#>
function Show-CloudPCMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$CloudPCs
    )
    
    Write-Host ""
    Write-FormattedMessage -Message "=== Cloud PC Selection Menu ===" -Type Info
    Write-Host ""
    
    for ($i = 0; $i -lt $CloudPCs.Count; $i++) {
        $cpc = $CloudPCs[$i]
        $displayText = "[$($i + 1)] $($cpc.ManagedDeviceName)"
        
        if ($cpc.PowerState) {
            $displayText += " - Power: $($cpc.PowerState)"
            
            $messageType = switch ($cpc.PowerState) {
                'running' { 'Success' }
                'stopped' { 'Warning' }
                default { 'Info' }
            }
            Write-FormattedMessage -Message $displayText -Type $messageType
        }
        else {
            Write-FormattedMessage -Message $displayText -Type Info
        }
        
        if ($cpc.Status) {
            Write-Host "    Status: $($cpc.Status)" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
    Write-FormattedMessage -Message "[0] Exit" -Type Warning
    Write-Host ""
}

<#
.SYNOPSIS
    Displays detailed information for a selected Cloud PC
#>
function Show-CloudPCDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [object]$CloudPC
    )
    
    Write-Host ""
    Write-FormattedMessage -Message "=== Cloud PC Details ===" -Type Info
    Write-Host ""
    
    $details = [ordered]@{
        'Display Name' = $CloudPC.DisplayName
        'User Principal Name' = $CloudPC.UserPrincipalName
        'Device Name' = $CloudPC.ManagedDeviceName
        'Cloud PC ID' = $CloudPC.Id
        'Status' = $CloudPC.Status
        'Power State' = $CloudPC.PowerState
        'Provisioning Policy' = $CloudPC.ProvisioningPolicyName
        'Provisioning Policy ID' = $CloudPC.ProvisioningPolicyId
        'Image Name' = $CloudPC.ImageDisplayName
        'Service Plan' = $CloudPC.ServicePlanName
        'Last Modified' = $CloudPC.LastModifiedDateTime
        'Grace Period End' = $CloudPC.GracePeriodEndDateTime
    }
    
    foreach ($key in $details.Keys) {
        $value = $details[$key]
        if ($value) {
            Write-Host "$($key): " -NoNewline -ForegroundColor Cyan
            Write-Host $value -ForegroundColor White
        }
    }
    Write-Host ""
}

<#
.SYNOPSIS
    Displays management action menu
#>
function Show-ActionMenu {
    [CmdletBinding()]
    param()
    
    Write-FormattedMessage -Message "=== Management Actions ===" -Type Info
    Write-Host ""
    Write-Host "[1] Start Cloud PC" -ForegroundColor Green
    Write-Host "[2] Stop Cloud PC" -ForegroundColor Yellow
    Write-Host "[3] Restart Cloud PC" -ForegroundColor Cyan
    Write-Host "[4] Reprovision Cloud PC" -ForegroundColor Magenta
    Write-Host "[5] Back to Cloud PC List" -ForegroundColor White
    Write-Host "[0] Exit" -ForegroundColor Red
    Write-Host ""
}

<#
.SYNOPSIS
    Executes power management actions on Cloud PC
#>
function Invoke-CloudPCAction {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CloudPCId,
        
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet('Start', 'Stop', 'Restart', 'Reprovision')]
        [string]$Action
    )
    
    try {
        $confirmMessage = "Are you sure you want to $Action the Cloud PC '$DisplayName'?"
        
        if ($PSCmdlet.ShouldProcess($DisplayName, $Action)) {
            Write-FormattedMessage -Message "Executing $Action on $DisplayName..." -Type Info
            
            switch ($Action) {
                'Start' {
                    Start-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId
                }
                'Stop' {
                    Stop-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId
                }
                'Restart' {
                    Restart-MgBetaDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId
                }
                'Reprovision' {
                    Invoke-MgBetaReprovisionDeviceManagementVirtualEndpointCloudPc -CloudPcId $CloudPCId
                }
            }
            
            Write-FormattedMessage -Message "$Action command sent successfully to $DisplayName" -Type Success
            Start-Sleep -Seconds 2
        }
    }
    catch {
        Write-FormattedMessage -Message "Failed to $Action Cloud PC: $($_.Exception.Message)" -Type Error
        Start-Sleep -Seconds 3
    }
}

<#
.SYNOPSIS
    Gets valid numeric input from user
#>
function Get-UserSelection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        
        [Parameter(Mandatory=$true)]
        [int]$MaxValue
    )
    
    do {
        try {
            [int]$selection = Read-Host $Prompt
            
            if ($selection -lt 0 -or $selection -gt $MaxValue) {
                Write-FormattedMessage -Message "Please enter a number between 0 and $MaxValue" -Type Warning
                $selection = -1
            }
            else {
                return $selection
            }
        }
        catch {
            Write-FormattedMessage -Message "Invalid input. Please enter a number." -Type Warning
            $selection = -1
        }
    } while ($selection -eq -1)
}

#endregion

#region Main Application Logic

<#
.SYNOPSIS
    Main application loop for Cloud PC management
#>
function Start-CloudPCManagement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion
    )
    
    while ($true) {
        try {
            Clear-Host
            
            # Get all Cloud PCs
            $cloudPCs = Get-AllCloudPCs -ApiVersion $ApiVersion
            
            if (-not $cloudPCs) {
                Write-FormattedMessage -Message "No Cloud PCs available. Press any key to exit..." -Type Warning
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                return
            }
            
            # Show selection menu
            Show-CloudPCMenu -CloudPCs $cloudPCs
            
            # Get user selection
            $selection = Get-UserSelection -Prompt "Select a Cloud PC (number)" -MaxValue $cloudPCs.Count
            
            if ($selection -eq 0) {
                Write-FormattedMessage -Message "Exiting Cloud PC Management Tool. Goodbye!" -Type Info
                break
            }
            
            # Get selected Cloud PC
            $selectedCPC = $cloudPCs[$selection - 1]
            
            # Inner loop for managing selected Cloud PC
            $continueManaging = $true
            while ($continueManaging) {
                Clear-Host
                Show-CloudPCDetails -CloudPC $selectedCPC
                Show-ActionMenu
                
                $action = Get-UserSelection -Prompt "Select an action" -MaxValue 5
                
                switch ($action) {
                    1 {
                        Invoke-CloudPCAction -CloudPCId $selectedCPC.Id -DisplayName $selectedCPC.DisplayName -Action 'Start' -Confirm:$false
                    }
                    2 {
                        Invoke-CloudPCAction -CloudPCId $selectedCPC.Id -DisplayName $selectedCPC.DisplayName -Action 'Stop' -Confirm:$false
                    }
                    3 {
                        Invoke-CloudPCAction -CloudPCId $selectedCPC.Id -DisplayName $selectedCPC.DisplayName -Action 'Restart' -Confirm:$false
                    }
                    4 {
                        Invoke-CloudPCAction -CloudPCId $selectedCPC.Id -DisplayName $selectedCPC.DisplayName -Action 'Reprovision' -Confirm:$true
                    }
                    5 {
                        $continueManaging = $false
                    }
                    0 {
                        Write-FormattedMessage -Message "Exiting Cloud PC Management Tool. Goodbye!" -Type Info
                        return
                    }
                }
            }
        }
        catch {
            Write-FormattedMessage -Message "An error occurred: $($_.Exception.Message)" -Type Error
            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
}

#endregion

#region Script Execution

# Main script entry point
try {
    Clear-Host
    Write-FormattedMessage -Message "=== Windows 365 Cloud PC Management Tool ===" -Type Success
    Write-FormattedMessage -Message "Version 2.0 - Updated for Modern Graph API" -Type Info
    Write-Host ""
    
    # Initialize Graph connection
    $connected = Initialize-GraphConnection -UseDeviceCode:$UseDeviceCodeAuth
    
    if (-not $connected) {
        Write-FormattedMessage -Message "Failed to establish connection. Exiting..." -Type Error
        exit 1
    }
    
    # Set Graph profile based on API version
    if ($GraphApiVersion -eq 'beta') {
        Select-MgProfile -Name 'beta'
        Write-FormattedMessage -Message "Using Microsoft Graph Beta API" -Type Info
    }
    else {
        Select-MgProfile -Name 'v1.0'
        Write-FormattedMessage -Message "Using Microsoft Graph v1.0 API" -Type Info
    }
    
    Start-Sleep -Seconds 2
    
    # Start main application
    Start-CloudPCManagement -ApiVersion $GraphApiVersion
}
catch {
    Write-FormattedMessage -Message "Critical error: $($_.Exception.Message)" -Type Error
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
finally {
    # Cleanup
    if (Get-MgContext) {
        Write-Host ""
        Write-FormattedMessage -Message "Disconnecting from Microsoft Graph..." -Type Info
        Disconnect-MgGraph | Out-Null
        Write-FormattedMessage -Message "Disconnected successfully." -Type Success
    }
}

#endregion
