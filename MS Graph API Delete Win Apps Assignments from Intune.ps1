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
    - Graph API permissions (Delegated):
        ·      DeviceManagementApps.ReadWrite.All
        ·      Group.Read.All
    - Modules installed in the Automation Account:
        • Microsoft.Graph.Authentication

#>


#Tenant Name
$Tenant = "MyTenant.onmicrosoft.com"
$Groups = @("INTUNE-MyWindows10", "INTUNE-MyWindows11") #Entra ID groups for which assignments need to be deleted

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


ForEach($WinApp in $WinApps){
  foreach($AppAssignment in $WinApp.assignments){
    if ($ADgroups.id -contains $AppAssignment.target.groupId) {
      Write-Host "Assignment found for App: $($WinApp.displayName)"
      Write-Host "Assignment group Id: $($AppAssignment.target.groupId)"
      $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($WinApp.id)/assignments/$($AppAssignment.id)"
      
      try {
        Connect-MgGraph -TenantId $Tenant -NoWelcome
        Invoke-MgGraphRequest -uri $uri -method DELETE -ContentType "application/json" 
        Write-Host "Action: Deleted" -ForegroundColor Green
        Write-Host ""
      }
      catch {
        $_.Exception.Message
        Write-Host "Action: ERROR" -ForegroundColor Red
        Write-Host ""
      }
      
    }
  }
}



