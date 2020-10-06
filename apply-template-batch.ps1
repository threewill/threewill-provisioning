
<#     
    .SYNOPSIS
    Take a batch of sites defined in a sites file and apply templates

    .DESCRIPTION
    This script will loop through the sites defined in the provided site file and construct urls from the config file and use this
    to call the apply-template.ps1 to apply the appropriate template to the site

    .PARAMETER configFile
    Relative path of the configuration file to use (e.g. config\prod.json)
    .PARAMETER sitesFile
    Relative path of a csv file containing a collection of site information

    .EXAMPLE
    .\apply-template-batch.ps1 -configFile config\contoso.json -sitesFile config\test-template-sites.csv

    .NOTES
    - Dependencies: 
        SharePointPnPPowerShellOnline cmdlets, version 3.12.1908.1 or higher (August 2019 Intermediate Release 1)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$configFile,
    [Parameter(Mandatory = $true)]
    [string]$sitesFile    
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"
$scriptTime = [System.Diagnostics.Stopwatch]::StartNew()

Write-Output "".PadRight(50, "=")
Write-Output "apply-template-batch.ps1"
Write-Output "  params: configFile:$($configFile),sitesFile:$($sitesFile)"
Write-Output "".PadRight(50, "=")

# Get configuration
$config = Get-Content $configFile | Out-String | ConvertFrom-Json

# Write out configuration information
Write-Output "Start Config Values ".PadRight(50, "*")
Write-Output $config
Write-Output "End Config Values ".PadRight(50, "*")

$global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."

if ($null -ne $global:cred) {

    $sitesData = Import-CSV (Join-Path $script:PSScriptRoot $sitesFile)

    foreach ($line in $sitesData) {
        if ($line.PlaceID.StartsWith("#")) {
            # This CSV line is "commented out" - so skip it
            continue
        }                

        if ($line.SiteType -ne "CommunicationSite") {
            Write-Warning "Skipping $($line.Site) because it is not a CommunicationSite SiteType"
        }
        else {
            $siteUrl = "$($config.rootSiteUrl)/sites/$($line.Site)"

            Write-Output "Start Invoke-Expression".PadRight(50, "*")
            Write-Output ".\apply-template.ps1"
            Write-Output "-configFile '$($configFile)'"
            Write-Output "-siteUrl '$($siteUrl)'"
            Write-Output "-skipGetCredentials -batchMode"
            Write-Output ""
            
            Invoke-Expression ".\apply-template.ps1 -configFile '$($configFile)' -siteUrl '$($siteUrl)' -skipGetCredentials -batchMode"

            Write-Output "End Invoke-Expression".PadRight(50, "*")
            Write-Output ""
        }
    }
}

Write-Output "".PadRight(50, "=")
Write-Output "Execution Time: [$(get-date)] [$($scriptTime.Elapsed.ToString())]"