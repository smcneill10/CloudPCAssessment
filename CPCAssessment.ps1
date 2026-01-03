
# https://techcommunity.microsoft.com/t5/windows-it-pro-blog/now-available-microsoft-graph-windows-365-apis/ba-p/4094550
# https://developer.microsoft.com/en-us/graph/changelog
# https://learn.microsoft.com/en-us/graph/use-the-api?context=graph%2Fapi%2Fbeta&view=graph-rest-beta#version

# https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell?view=powershell-7.3
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0

#Arm
# https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-on-arm?view=powershell-7.4

#Entra ID Powershell info
#https://techcommunity.microsoft.com/t5/microsoft-entra-blog/introducing-the-microsoft-entra-powershell-module/ba-p/4173546

#Display Preferences
$FGColor = "white"
$BKColor = "Black"
$BKColorBad = "Red"
#$BKColorGood = "Green"
$BKColorinfo = "black"

#Import the required modules - V1.0 versus Beta
Import-Module Microsoft.Graph.DeviceManagement.Administration -Force 
Import-Module Microsoft.Graph.beta.DeviceManagement.Administration -Force 
#import-Module -Name Microsoft.Graph -Force -AllowClobber

#Connect to CloudPC Graph API 
Connect-MgGraph -Scopes "CloudPC.ReadWrite.All, User.Read.All","Group.Read.All, CloudPC.read.all"


#Gathers the connection info, comment out the Clear-Host line below to see this info, helps with connectivity issues
Write-host "Here is the connection information used:" #-BackgroundColor $BKColorInfo -ForegroundColor $FGColor
Get-MgContext
#Clear-Host

#Function to gather CPC info and allow for mgmt
Function Get-CloudPCData  
    {
    write-host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
    Write-host "Enter 1 for API V1.0 and 2 for API Beta" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor

    [int]$APIVersion = Read-Host "Enter your selection"
    If ($APIVersion -eq 1) 
        {Write-Host "Using API V1.0" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        #get all CPCs using V1.0 API version
        $CPCs = Get-MgDeviceManagementVirtualEndpointCloudPc -Property DisplayName, UserPrincipalName, ManagedDeviceName, ID, ProvisioningPolicyId, ProvisioningPolicyName, ImageDisplayName, ServicePlanName
        }
    Else 
        {Write-Host "Using API Beta" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        # get all CPCs using Beta API version
        $CPCs = Get-MgBetaDeviceManagementVirtualEndpointCloudPc -Property DisplayName, UserPrincipalName, ManagedDeviceName, ID, ProvisioningPolicyId, ProvisioningPolicyName, ImageDisplayName, ServicePlanName, Status, PowerState
        }
    $Counter = 0
    # cycle thru all CPCs and display info
    foreach ($CPC in $CPCs)
    {
        $counter++
        $RunningStatus = "Running"
        If ($null -ne $CPC.PowerState)
            {
                $runningStatus = $CPC.Powerstate

                Write-Host "Select" $Counter "for" $CPC.ManagedDeviceName "    " $runningStatus -BackgroundColor $BKColorInfo
            }
        Else
            {
                write-Host "Select" $Counter "for" $CPC.ManagedDeviceName  -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
            }
    }
  
    Write-host "Select 0 to exit" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
    Write-Host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor

    #get the selection for detailed info for CPC
    [int]$Selection1 = Read-Host "enter number for more info and to Manage a CPC " 
    If ($Selection1 -eq 0) {Write-Host "Thanks and See Ya" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor; Break} 
    If ($Selection1 -gt $counter) {Write-host ""; Write-host "Out of band selection, please select again" -ForegroundColor $FGColor -backgroundcolor $BKColorBad; Get-CloudPCData}
    $choosenCPC = $selection1 -1

        #Display detailed info for selected CPC
        Write-Host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "Cloud PC Display Name:" $CPCs[$choosenCPC].DisplayName -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC User Name:" $CPCs[$choosenCPC].UserPrincipalName -ForegroundColor $FGColor -BackgroundColor $BKColor
        write-host "CLoud PC NETBIOS Name:" $CPCs[$choosenCPC].ManagedDeviceName -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC ID:"  $CPCs[$choosenCPC].Id -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC Status:"  $CPCs[$choosenCPC].Status -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC Provisioning Policy ID:"$CPCs[$choosenCPC].ProvisioningPolicyId -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC Provisioning Policy Name:"$CPCs[$choosenCPC].ProvisioningPolicyName -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC Provisioning Policy Image Name:"$CPCs[$choosenCPC].ImageDisplayName -ForegroundColor $FGColor -BackgroundColor $BKColor
        Write-Host "Cloud PC Sevice Plan Name:"$CPCs[$choosenCPC].ServicePlanName -ForegroundColor $FGColor -BackgroundColor $BKColor
        If ($null -ne $CPCs[$choosenCPC].PowerState )
        {Write-Host "Cloud PC Power State:"$CPCs[$choosenCPC].PowerState -ForegroundColor $FGColor -BackgroundColor $BKColor}
        Write-Host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor

       #Display optional actions for selected CPC
        Write-Host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "Optional Action Menu" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "1" "Start" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "2" "Stop" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "3" "Restart" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "4" "Connectivity History" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "5" "Back" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "6" "Exit" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
        Write-Host "" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor

        



    [int]$Selection2 = Read-Host "Enter your selection"
    #Switch for the optional actions
    Switch ($Selection2)
    {
    1 {Start-MgDeviceManagementVirtualEndpointCloudPcOn -CloudPCId $CPCs[$choosenCPC].Id}
    1 {write-host 'Starting'  $CPCs[$choosenCPC].DisplayName}
    1 {Get-CloudPCData}
    2 { Start-MgDeviceManagementVirtualEndpointCloudPcOff -CloudPCId $CPCs[$choosenCPC].Id }
    2 {write-host 'Stopping ' $CPCs[$choosenCPC].DisplayName}
    2 {Get-CloudPCData}
    3 {Restart-MgDeviceManagementVirtualEndpointCloudPc -CloudPcId $CPCs[$choosenCPC].Id }
    3 {write-host 'Re-Starting ' $CPCs[$choosenCPC].DisplayName}
    3 {Get-CloudPCData}
    4 {Get-CPCConnectHistory $CPCs[$choosenCPC].DisplayName $CPCs[$choosenCPC].Id $CPCs[$choosenCPC].ManagedDeviceName}
    #4 {Clear-Host}
    4 {Get-CloudPCData}
    5 {Clear-host}
    5 {Get-CloudPCData}
    6 {Write-Host 'See Ya'}
    6 {break}
    Default {Get-CloudPCData }
    }
}




#start of script

Get-CloudPCData