
# Requires: PowerShell Universal Dashboard, Microsoft.Graph modules

Import-Module UniversalDashboard
Import-Module Microsoft.Graph

Start-UDDashboard -Port 10000 -Dashboard (
    New-UDDashboard -Title "Windows 365 Cloud PC Request" -Content {
        New-UDForm -Title "Request a Cloud PC" -Content {
            New-UDTextbox -Id "userPrincipalName" -Label "Your Email Address"
            New-UDSelect -Id "region" -Label "Select Region" -Options @(
                New-UDSelectOption -Name "US East" -Value "us-east"
                New-UDSelectOption -Name "Europe West" -Value "eu-west"
                # Add more regions as needed
            )
            New-UDSelect -Id "specs" -Label "Select Cloud PC Specs" -Options @(
                New-UDSelectOption -Name "2 vCPU, 8GB RAM, 128GB" -Value "2vCPU"
                New-UDSelectOption -Name "4 vCPU, 16GB RAM, 128GB" -Value "4vCPU"
            )
        } -OnSubmit {
            param($userPrincipalName, $region, $specs)

            # Licensing check
            $user = Get-MgUser -UserId $userPrincipalName
            $hasLicense = $user.AssignedLicenses | Where-Object { $_.SkuId -eq "<W365_SKU_GUID>" }
            if (-not $hasLicense) {
                Show-UDToast -Message "You do not have the required Windows 365 license." -Duration 5000 -Position topCenter
                return
            }

            # Provisioning logic (pseudo-code, replace with actual API calls)
            $provisionResult = "Success" # Replace with actual provisioning logic

            if ($provisionResult -eq "Success") {
                Show-UDToast -Message "Cloud PC provisioned successfully in $region with $specs." -Duration 5000 -Position topCenter
            } else {
                Show-UDToast -Message "Provisioning failed: $provisionResult" -Duration 5000 -Position topCenter
            }
        }
    }
)
