# Need to have PowerShell 7 or higher https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell?view=powershell-7.3




#Connect to CloudPC Graph API 
Connect-MgGraph -Scopes "CloudPC.ReadWrite.All, User.Read.All","Group.Read.All, CloudPC.read.all"
#Set Graph API to Beta
#Select-MgProfile Beta

#Gathers the connection info, comment out the Clear-Host line below to see this info, helps with connectivity issues
Write-host "Here is the connection information used:" -BackgroundColor $BKColorInfo -ForegroundColor $FGColor
Get-MgContext
#Clear-Host

