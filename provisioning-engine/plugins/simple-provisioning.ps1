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
    [string]$ConfigFile,
    [string]$SiteType
)

Write-Log "[$FullSiteUrl] Simple Provisioning plugin started" -WriteToHost
Write-Log "[$FullSiteUrl] Site Type passed in: $SiteType" -WriteToHost

# Get configuration
$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

# Establish custom environment variables from our config that was passed in
$variables = $config.environmentVariables
$SpokeSiteTemplate = ""

foreach ($var in $variables) {
  if(Get-Member -inputobject $var -name "SpokeSiteTemplate" -Membertype Properties){
    #Property exists
    $SpokeSiteTemplate = $var.SpokeSiteTemplate
    Write-Log "Found Spoke Template: $SpokeSiteTemplate" -WriteToHost
  }
}

Write-Log "Since this is a provisioning script (not pre- or post-), site creation is an assumed responsibility of this script.  Creating first..." -WriteToHost

$adminConn = Connect-PnPOnline -Url $config.adminSiteUrl -Credentials $global:cred -ReturnConnection -ErrorAction Stop

Write-Log "New-PnPSite -Type CommunicationSite -Title $($SiteTitle) -Url $($siteUrl)"
$newSiteUrl = New-PnPSite -Type CommunicationSite `
    -Title $SiteTitle `
    -Url $siteUrl `
    -Description $SiteTitle `
    -SiteDesign "Topic" `
    -Connection $adminConn

Disconnect-PnPOnline

Write-Log "Provisioning called (intranet): Tenant Url $TenantUrl | SitePath $SitePath | SiteTitle $SiteTitle | FullSiteUrl $FullSiteUrl" -WriteToHost

# Connect to target site
$connection = Connect-PnPOnline -Url $newSiteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop

Write-Log "[$FullSiteUrl] Invoking PnP template... $PSScriptRoot/$SpokeSiteTemplate" -WriteToHost
# Ensure the template file exists
# if (Test-Path -Path "$PSScriptRoot/$SpokeSiteTemplate")
# {
  
  # We have to handle this template in a special way due to Managed Metadata
  # See discussion here: https://github.com/pnp/PnP-PowerShell/issues/1180#issuecomment-583447230

  # Invoke only the fields
  #Invoke-PnPSiteTemplate -Path "$PSScriptRoot/$SpokeSiteTemplate" -Connection $connection
  Write-Log "Feaux site provisioning occurs here" -WriteToHost
# }
# else {
#   Write-Log "Could not find provisioning script '$PSScriptRoot/$SpokeSiteTemplate'" -Level Error
# }

Write-Log "[$FullSiteUrl] Simple Provisioning plugin completed" -WriteToHost

# Disconnect the site connection
Disconnect-PnPOnline -Connection $connection
