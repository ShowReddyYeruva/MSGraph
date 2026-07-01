
##########################Intune portal token#####################
$Token = "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6............"  ####### Copy token from Intune console#########


$authHeader = @{
    'Content-Type'  = 'application/json'
    'Authorization' = $Token
    'ExpiresOn'     = ''
}



########################Get All Applications#############
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
    $AppBatch = Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Get -UseBasicParsing -ContentType "application/json"
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


###########Filter for Win apps######################
$WinApps = $AllApps | Where-Object {($_.'@odata.type' -like '#microsoft.graph.win*') -or ($_.'@odata.type' -like '#microsoft.graph.officeSuiteApp*')}
$WinApps.count


############################Create report with required properties and Entra ID group name###########


$Report = @()
foreach($WinApp in $WinApps){
  $WinAppAssignments = $WinApp.assignments
  foreach($WinAppAssignment in $WinAppAssignments){
    $Result = $null
    $Result = New-Object PsObject -property @{'displayName' = $WinApp.displayName}
    $Result | Add-Member -MemberType NoteProperty -Name 'displayVersion' -Value $WinApp.displayVersion
    $Result | Add-Member -MemberType NoteProperty -Name 'publisher' -Value $WinApp.publisher

    If($WinApp.notes -like 'PmpAppId*'){
      $Result | Add-Member -MemberType NoteProperty -Name 'CreatedBy' -Value 'PatchMyPC'
    }
    else {
      $Result | Add-Member -MemberType NoteProperty -Name 'CreatedBy' -Value 'inHouseTeam'
    }

    $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentIntent' -Value $WinAppAssignment.intent

    If($WinAppAssignment.target.'@odata.type' -like '*exclusionGroupAssignment*'){
      $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentIntentMode' -Value 'Excluded'
    }
    else {
      $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentIntentMode' -Value 'Included'
    }

    If($WinAppAssignment.target.'@odata.type' -like '*allDevicesAssignment*'){
      $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentGroup' -Value 'All devices'
    }
    elseif ($WinAppAssignment.target.'@odata.type' -like '*allLicensedUsersAssignment*') {
      $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentGroup' -Value 'All users'
    }
    else {
      $ADgroupsName = $null
      $graphApiVersion = "beta"
      $Resource = "groups"
      $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$($WinAppAssignment.target.groupId)"
      $ADgroupsName = Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Get -UseBasicParsing -ContentType "application/json"

      $Result | Add-Member -MemberType NoteProperty -Name 'AssignmentGroup' -Value $ADgroupsName.displayName
    }
    $Report += $Result
    $Report.Count
    $Result
  }

}


$DateTime = get-date -format yyyy-MM-ddTHH-mm-ss-ff
$FileName = "WindoiwsAppsAssignments_Report_" + $DateTime
$Report | Export-Csv -NoTypeInformation -Path "C:\Intune\Reports\$FileName.csv"




