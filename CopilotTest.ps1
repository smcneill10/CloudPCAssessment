# Install the Microsoft Graph PowerShell module
Install-Module -Name Microsoft.Graph

# Import the module
Import-Module -Name Microsoft.Graph

# Connect to Microsoft Graph API
Connect-MgGraph -ClientId 'YOUR_CLIENT_ID' -TenantId 'YOUR_TENANT_ID' -ClientSecret 'YOUR_CLIENT_SECRET'

# Example: Get all Cloud PCs
$cloudPCs = Get-MgCloudPc

# Example: Get details of a specific Cloud PC
$cloudPC = Get-MgCloudPc -Id 'YOUR_CLOUD_PC_ID'