param(
    [Parameter(Mandatory = $true)]
    [string]$TenantUrl,
    [Parameter(Mandatory = $true)]
    [string]$SitePath,
    [Parameter(Mandatory = $true)]
    [string]$SiteTitle,
    [Parameter(Mandatory = $true)]
    [string]$FullSiteUrl,
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile
)

Write-Host "[$FullSiteUrl]: Intranet Provisioning plugin started"

# Get configuration
$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

# Establish custom environment variables from our config that was passed in
$variables = $config.environmentVariables
$hub = ""
$SpokeSiteTemplate = ""

foreach ($var in $variables) {
  if(Get-Member -inputobject $var -name "Hub" -Membertype Properties){
    #Property exists
    $hub = $var.hub
    Write-Host "Found Hub: $hub"
  }
  if(Get-Member -inputobject $var -name "SpokeSiteTemplate" -Membertype Properties){
    #Property exists
    $SpokeSiteTemplate = $var.SpokeSiteTemplate
    Write-Host "Found Spoke Template: $SpokeSiteTemplate"
  }
}

Write-Host "Provisioning called (intranet): Tenant Url" $TenantUrl "| SitePath" $SitePath "| SiteTitle" $SiteTitle "| FullSiteUrl" $FullSiteUrl

# Connect to target site
$connection = Connect-PnPOnline -Url $FullSiteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop

Write-Host "[$FullSiteUrl]: Invoking PnP template... $PSScriptRoot/site-template.xml"
# Ensure the template file exists
if (Test-Path -Path "$PSScriptRoot/$SpokeSiteTemplate")
{
  
  # We have to handle this template in a special way due to Managed Metadata
  # See discussion here: https://github.com/pnp/PnP-PowerShell/issues/1180#issuecomment-583447230

  # Invoke only the fields
  Invoke-PnPSiteTemplate -Path "$PSScriptRoot/$SpokeSiteTemplate" -Connection $connection -Handlers Fields

  # Ensure that the Taxonomy feature is enabled
  Enable-PnPFeature -Identity 73ef14b1-13a9-416b-a9b5-ececa2b0604c -Connection $connection -Scope Site

  # Wait for the taxonomy field to appear
  Write-Host "[$FullSiteUrl]: Ensuring taxonomy is avaiable"
  $timeElapsed = 0
  $fieldExists =  Get-PnPField -Identity "TermStoreCategories" -Connection $connection -ErrorAction SilentlyContinue
  while($null -eq $fieldExists){
    # Exit after 5 minutes
    $timeElapsed += 5
    if ($timeElapsed -ge 300)
    {
      Write-Error "Timeout occurred waiting on taxonomy provisioning"
    }

    Write-Host "Waiting..." 
    Start-Sleep -Seconds 5

    $fieldExists =  Get-PnPField -Identity "TermStoreCategories" -Connection $connection -ErrorAction SilentlyContinue
  }
  Write-Host "Taxonomy field found"

  # Invoke the remainder of the teamplate
  Invoke-PnPSiteTemplate -Path "$PSScriptRoot/$SpokeSiteTemplate" -Connection $connection -ClearNavigation -ExcludeHandlers Fields

}
else {
  Write-Error "Could not find provisioning script '$PSScriptRoot/$SpokeSiteTemplate'"
}


#### Post provisioning tasks ####
Write-Host "[$FullSiteUrl]: Running post-provisioning tasks..."
Write-Host "[$FullSiteUrl]: Customizing Contribute Permission Level and applying to Site Members"
# Update the contribute permission to remove delete permission
$role = Set-PnPRoleDefinition -Identity "Contribute" -Clear DeleteListItems, DeleteVersions -Description "Can view, add, and update list items and documents." -Connection $connection
try {
  # Update the members group to use the Contribute Role
  # This will throw an error if the script is run a second time
  $memberGroup = Get-PnPGroup -AssociatedMemberGroup
  $perm = Set-PnPGroupPermissions -Identity $memberGroup.Id -RemoveRole "Edit" -AddRole "Contribute" -Connection $connection
}
catch {
  Write-Host "Could not update permissions for '$SiteTitle Members'"
  Write-Host $_
}

### Break inheritance on the Home page and make it so that Members only have Read permission level
Write-Host "[$FullSiteUrl]: Breaking Home page permissions and giving Members Read access"
$homePage = Get-PnPClientSidePage "Home.aspx"
$homePageItem = Get-PnPListItem -List "Site Pages" -Id $homePage.PageId
$homePageItem.BreakRoleInheritance($True,$False)
$homePageItem.Update()
$memberGroup = Get-PnPGroup -AssociatedMemberGroup
Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.PageId -Group $memberGroup.Id -RemoveRole "Contribute"
Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.PageId -Group $memberGroup.Id -AddRole "Read"

#Associate the site to the Hub
Write-Host "[$FullSiteUrl]: Associating to Community Hub ($hub)"
Add-PnPHubSiteAssociation -Site $FullSiteUrl -HubSite "$hub" -Connection $connection

Write-Host "[$FullSiteUrl]: Intranet Provisioning plugin completed"

# Disconnect the site connection
Disconnect-PnPOnline -Connection $connection
