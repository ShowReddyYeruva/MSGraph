<#
.SYNOPSIS
    Approve Intune multi admin approval requests in bulk.

.DESCRIPTION
    This PS script fetch Intune multi admin approval requests and bulk approve for spacific justification.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - $Tenant : Azure AD tenant name
    - $ActivityJustification : Enter justification of the activity for bulk approval
    - $ApprovalJustification : Approval justification
    - Graph API permissions (Delegated):
        ·      DeviceManagementApps.ReadWrite.All
        ·      Group.Read.All
    - Modules installed in the Automation Account:
        • Microsoft.Graph.Authentication

.DISCLAIMER
    - This script is provided "AS IS" without warranty of any kind.
    - Test thoroughly in a non-production environment before use.
    - The author is not responsible for any impact, data loss, or service disruption resulting from the execution of this script. Use at your own risk
#>

$Tenant = "MyTenant.onmicrosoft.com"
$ActivityJustification = "Test"
$ApprovalJustification = "Approved"


Write-Host ("Checking for Microsoft.Graph.Authentication module...")

#Get Installed AzureAD Module
$MgGraphRequestModule = Get-Module -Name "Microsoft.Graph.Authentication" -ListAvailable
  
#Check if it got one for Microsoft.Graph.Authentication module
if ($MgGraphRequestModule -eq $null)
{
    #if not Ask user to install 'Microsoft.Graph.Authentication' module
    write-host
    write-host ("Microsoft.Graph.Authentication Powershell module not installed...")
    write-host ("Install by running 'Install-Module Microsoft.Graph.Authentication from an elevated PowerShell prompt")
    write-host ("Script can't continue...")
    write-host
    exit
}
else {
    write-host ("Microsoft.Graph.Authentication Powershell module installed!!")
}

Import-Module Microsoft.Graph.Authentication -Force

Connect-MgGraph -TenantId $Tenant -NoWelcome


Write-Host "Fectching MAA requests for justification: $ActivityJustification" -ForegroundColor Green

$graphApiVersion = "beta"
$Resource = "deviceManagement/operationApprovalRequests"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)`?`$filter=status eq 'needsApproval' and requestJustification eq '$ActivityJustification'"
$MAA_requests = (Invoke-MgGraphRequest -uri $uri -method GET -ContentType "application/json").value



Write-Host "$($MAA_requests.count) requests found." -ForegroundColor Green
Write-Host ""

Write-Host "Approval in progress..."

$JSON_Output = @"
    {
    "justification":$ApprovalJustification
    }
"@

foreach($MAA_request in $MAA_requests){
  $graphApiVersion = "beta"
  $Resource = "deviceManagement/operationApprovalRequests"
  $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$($MAA_request.id)/approve"
  Write-Host "Processing: $($MAA_request.payloadName)"
  Write-Host "Resource type: $($MAA_request.requiredApprovalRequestPolicies)"
  Write-Host "Operation: $($MAA_request.payloadOperation)"
  Write-Host "Business justification: $($MAA_request.requestJustification)"
  Write-Host "Requested by: $($MAA_request.requestor.user.displayName)"
  
  try{
    Connect-MgGraph -TenantId $Tenant -NoWelcome
    Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Post -Body $JSON_Output -ContentType "application/json"
    Write-Host "Status: Success" -ForegroundColor Green
  }
  catch {
    $Fail = $_.Exception.Message
    Write-Host "Status: Error" -ForegroundColor Red
    Write-Host "Error message: $Fail" -ForegroundColor Red
  }
  Write-Host ""
}

