# Need to have PowerShell 7 or higher https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell?view=powershell-7.3
# https://techcommunity.microsoft.com/t5/windows-it-pro-blog/now-available-microsoft-graph-windows-365-apis/ba-p/4094550
# https://developer.microsoft.com/en-us/graph/changelog
# https://learn.microsoft.com/en-us/graph/use-the-api?context=graph%2Fapi%2Fbeta&view=graph-rest-beta#version

#Install the required modules
Install-Module -Name Microsoft.Graph.Authentication -Force
Install-Module -Name Microsoft.Graph.PowerShell -Force
install-Module -Name Microsoft.Graph.DeviceManagement.Administration -Force

#Connect to CloudPC Graph API 
Connect-MgGraph -Scopes "CloudPC.ReadWrite.All, User.Read.All","Group.Read.All, CloudPC.read.all"
#Set Graph API to Beta or V1.0
#Select-MgProfile Beta

#Gathers the connection info, comment out the Clear-Host line below to see this info, helps with connectivity issues
Write-host "Here is the connection information used:" #-BackgroundColor $BKColorInfo -ForegroundColor $FGColor
Get-MgContext
#Clear-Host

