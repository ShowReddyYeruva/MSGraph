<#
.SYNOPSIS
    Delete Application assignments for spacific Entra ID groups.

.DESCRIPTION
    This PS script checks for Windows apps and it's assignments and delete assignments targeted to spacific Entra ID groups.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - $Tenant : Azure AD tenant name
    - $Groups : Entra ID groups for which assignments need to be deleted
    - $ApprovalJustification : Enter meaningfull justification of your activity
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


#Tenant Name
$Tenant = "MyTenant.onmicrosoft.com"
$Groups = @("INTU-Windows10", "INTU-Windows11") #Entra ID groups for which assignments need to be deleted
$ApprovalJustification = "Bulk app assignment delete..."


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



########################Get all Applications#########
#Version of Graph API
#$graphApiVersion = "v1.0"
$graphApiVersion = "beta"
#Resource URI
$Resource = "deviceAppManagement/mobileApps"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)`?`$expand=assignments"

$AllApps = @()
# Invoke REST method and fetch data until there are no pages left.
do
{
  
  
  try {
    Connect-MgGraph -TenantId $Tenant -NoWelcome
    $AppBatch = Invoke-MgGraphRequest -uri $uri -method GET -ContentType "application/json"
  }
  catch {
    $_.Exception.Message
    $Fail = $_.Exception.Message
    #$Device.deviceName
    Start-Sleep -Seconds 20
  }

  if ($AppBatch.value)
  {
    $AllApps += $AppBatch.value
  } 
  else
  {
    $AllApps += $AppBatch
  }
  
  $uri = $AppBatch.'@odata.nextlink'
} until (!($uri))
  

$AllApps.count

$WinApps = $AllApps | Where-Object {($_.'@odata.type' -like '#microsoft.graph.win*') -or ($_.'@odata.type' -like '#microsoft.graph.officeSuiteApp*')}
$WinApps.count


################Get Entra groupp IDs####################

$ADgroups = @()

foreach($Group in $Groups){
  $graphApiVersion = "beta"
  $Resource = "groups"
  $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)`?`$filter= displayName eq '$Group'"
  Connect-MgGraph -TenantId $Tenant -NoWelcome
  $ADgroups += (Invoke-MgGraphRequest -uri $uri -method GET -ContentType "application/json").value
}


####################Delete assignments if matches with provided group IDs###############

##Encode justification provided
$approvalJustificationBytes = [System.Text.Encoding]::UTF8.GetBytes($ApprovalJustification)
$approvalJustificationBase64 = [System.Convert]::ToBase64String($approvalJustificationBytes)


foreach($WinApp in $WinApps){
  
  Write-Host "Preocessing: $($WinApp.displayName)" -ForegroundColor Green
 
  $mobileAppAssignments = @()
  $foundAssignmentToRemove = $false
  foreach ($assignment in $WinApp.assignments) {
      if ($ADgroups.id -contains $assignment.target.groupId) {
          $foundAssignmentToRemove = $true
          Write-Host "   Group Assignment found." -ForegroundColor Yellow
          continue
      }
      
      $assignment.PSObject.Properties.Remove("id")
      $assignment | Add-Member -NotePropertyName "@odata.type" -NotePropertyValue "#microsoft.graph.mobileAppAssignment" -Force
      $mobileAppAssignments += $assignment
  }

  
  if ($foundAssignmentToRemove) {
      Write-Host "   Removing assignments" -ForegroundColor Yellow
      $payload = @{
          mobileAppAssignments = $mobileAppAssignments
      }

      $JSON_Output = $payload | ConvertTo-Json -Depth 5
      #$JSON_Output

      
      $customHeaders = @{
        'x-msft-approval-justification' = $approvalJustificationBase64
      }

     
      $Uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($WinApp.id)/assign"
      #$Uri
      
      try {
        Connect-MgGraph -TenantId $Tenant -NoWelcome
        Invoke-MgGraphRequest -uri $uri -Headers $customHeaders -method Post -Body $JSON_Output -ContentType "application/json"
      }
      catch {
        $_.Exception.Message
        $Fail = $_.Exception.Message
      }

  }
  else {
    Write-Host "   No assignments found to delete." -ForegroundColor Green
  }

Write-Host "" -ForegroundColor Green
Write-Host "" -ForegroundColor Green
}




