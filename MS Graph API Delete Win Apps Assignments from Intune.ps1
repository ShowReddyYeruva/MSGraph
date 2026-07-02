<#
.SYNOPSIS
    Delete Application assignments for spacific Entra ID groups.

.DESCRIPTION
    This PS script checks for Windows apps and it's assignments and delete assignments targeted to spacific Entra ID groups.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - $AppId : App registration client ID
    - $TenantId : Azure AD tenant ID
    - $client_secret : App registration client secret
    - $Groups : Entra ID groups for which assignments need to be deleted
    - Graph API permissions (Delegated):
        ·      DeviceManagementApps.ReadWrite.All
        ·      Group.Read.All

#>
###############################################
$AppId = #<Azure SPN Id>#
$TenantId = #<Azure tenant Id>#
$client_secret = #<Application secret>#
$Groups = @("INTU-Windows10", "INTU-Windows11") #Entra ID groups for which assignments need to be deleted
###############################################


function Get-AppAuthToken {
  
    $body = @{
        client_id     = $AppId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $client_secret
        grant_type    = "client_credentials"
    }
  
    try { 
      $tokenRequest = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop 
      $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
  
      If($token){
        $authToken = @{
          'Content-Type'='application\json'
          'Authorization'="Bearer $token"
        }
        
        return $authToken
      }
      else {
        throw "An error occured getting access token: $($_.Exception.Message)"
      }
  
  
    }
  
    catch { 
      throw $_.Exception.Message  
    }
  
}


$AppAuthToken = Get-AppAuthToken





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
    $AppBatch = Invoke-RestMethod -Uri $uri -Headers $AppAuthToken -Method Get -UseBasicParsing -ContentType "application/json"
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
  $ADgroups += (Invoke-RestMethod -Uri $uri -Headers $AppAuthToken -Method Get -UseBasicParsing -ContentType "application/json").value
}


####################Delete assignments if matches with provided group IDs###############

ForEach($WinApp in $WinApps){
  foreach($AppAssignment in $WinApp.assignments){
    if ($ADgroups.id -contains $AppAssignment.target.groupId) {
      Write-Host "Assignment found for App: $($WinApp.displayName)"
      Write-Host "Assignment group Id: $($AppAssignment.target.groupId)"
      $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($WinApp.id)/assignments/$($AppAssignment.id)"
      
      try {
        Invoke-RestMethod -Uri $uri -Headers $AppAuthToken -Method DELETE -UseBasicParsing -ContentType "application/json"
      }
      catch {
        $_.Exception.Message
      }
      Write-Host "Action: Deleted"
      Write-Host ""
    }
  }
}





