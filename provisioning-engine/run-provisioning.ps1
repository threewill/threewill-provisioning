<#     
    .SYNOPSIS
    Create sites and teams based on a csv file

    .DESCRIPTION
    This script will create sites and teams based on the provided csv file.  It depends on the create-entity.ps1
    and provision-site.ps1 script that is invoked when this loops through the rows in the csv and constructs the call

    .PARAMETER ConfigFile
    Relative path of the configuration file to use (e.g. config\coxautoinc.json)
    .PARAMETER SitesFile
    Relative path of a csv file containing a collection of site information
    .PARAMETER UseHistory
    If provided, will expect to check a status file (e.g. status\timestamp\status.csv) for the status of previously created or provisioned entities
    .PARAMETER StatusFile
    The path to the status file

    .EXAMPLE
    .\run-provisioning.ps1 -ConfigFile config\boinga.json -SitesFile config\test-template-sites.csv

    .NOTES
    - Dependencies: 
        SharePointPnPPowerShellOnline cmdlets, version 3.12.1908.1 or higher (August 2019 Intermediate Release 1)
        Teams cmdlets, version 1.01 or higher (Install-Module -Name MicrosoftTeams)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,
    [Parameter(Mandatory = $true)]
    [string]$SitesFile,
    [Parameter(ParameterSetName = "History")]
    [switch]$UseHistory,
    [Parameter(ParameterSetName = "History", Mandatory = $false)]
    [string]$StatusFile = $null,
    [Parameter(Mandatory=$false)]
    [ValidateSet("Interactive", "Console")]
    [string]$AuthMode="Interactive",
    [Parameter(Mandatory=$false)]
    [string]$ClientId
)


$ErrorActionPreference = "Stop"
Set-StrictMode -Version "3.0"
$scriptTime = [System.Diagnostics.Stopwatch]::StartNew()
$statusTime = (Get-Date).ToUniversalTime().ToString("yyyyMMddHHmmss")
$global:logFile = "./logging/$statusTime.log"

# Ensure logging folder exists
if (-not (Test-Path -Path "./logging"))
{
    $loggingFolder = New-Item -Path "./logging" -ItemType Directory
 
}

# Make the helper functions available to script
. "helpers/helper-functions.ps1"

Write-Log "".PadRight(50, "=") -WriteToHost
Write-Log "run-provisioning.ps1" -WriteToHost
Write-Log "  params: ConfigFile:$($ConfigFile),SitesFile:$($SitesFile)" -WriteToHost
Write-Log "".PadRight(50, "=") -WriteToHost -WriteNewLine

# Get configuration
$config = Get-Content $ConfigFile | Out-String | ConvertFrom-Json

# Write out configuration information
Write-Log "Start Config Values ".PadRight(50, "*") -WriteToHost
Write-Log $config -WriteToHost
Write-Log "End Config Values ".PadRight(50, "*") -WriteToHost

$global:cred = Get-Credential -Message "Please Provide Credentials with SharePoint Admin permission."
If($AuthMode -eq "Interactive")
{
    $global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
}
else{
    $global:tenantConn = Connect-PnPOnline -Url $config.adminSiteUrl -Credential $global:cred -ReturnConnection -ErrorAction Stop
}
$global:urls = @()

enum ProcessingStatus {
    Creating
    ErrorCreating
    CreatingComplete
    Provisioning
    ErrorProvisioning
    Complete
    Skipped
}

try
{

    $siteStatus = @{}
    # Check for an existing status file
    if ($false -ne $UseHistory)
    {
        Write-Log "Checking for previous status file"
        if ($StatusFile)
        {
            if (Test-Path -Path $StatusFile)
            {
                Write-Log "Found status file '$StatusFile'"
                Import-Csv $StatusFile | % { $siteStatus[$_.SiteUrl] = [ProcessingStatus]$_.Status }
            }
            else
            {
                Write-Log "Status file '$StatusFile' not found"
                if (-not $PSCmdlet.ShouldContinue("Do you want to proceed without previous status?", "Status file '$StatusFile' not found.")) {
                    # User does not want to continue, end process
                    Write-Log "User chose not to continue.  Ending process."
                    exit
                }
            }
        }
        else {
            # Find the latest history file
            $statusFolders = $null
            
            # Check for status folder
            if (Test-Path -Path "./status")
            {
                $statusFolders = Get-ChildItem "./status" -Directory | Select Name
            }

            if ($null -ne $statusFolders -and $statusFolders.Length -gt 0)
            {   
                $statusFolder = $statusFolders | Select-Object Name | Sort-Object -Descending | Select-Object -Last 1
                $statusFile = "./status/$($statusFolder.Name)/status.csv"
                if (Test-Path -Path $statusFile)
                {
                    Write-Log "Found status file '$statusFile'"
                    Import-Csv $statusFile | % { $siteStatus[$_.SiteUrl] = [ProcessingStatus]$_.Status }
                }
                else
                {
                    Write-Log "Status file '$statusFile' not found"

                    if (-not $PSCmdlet.ShouldContinue("Do you want to proceed without previous status?", "Status file '$statusFile' not found")) {
                        # User does not want to continue, end process
                        Write-Log "User chose not to continue.  Ending process."
                        exit
                    }
                }
            }
            else
            {
                Write-Log "No history found.  Running script without previous status"
            }
        }
    }


    if ($null -ne $global:cred)
    {
        $sitesData = Import-CSV (Join-Path $script:PSScriptRoot $SitesFile) | Where { -not $_.PlaceID.StartsWith("#") }
       
        $nonTeamsItems = $sitesData | Where { $_.EntityType -ne "MSTeam" }
        if ($null -ne $nonTeamsItems -and $nonTeamsItems.Length -gt 0)
        {
            # If we have any sites to provision that are not MSTeam entities, connect to SharePoint
            Write-Log "Connect-PnpOnline" -WriteToHost
            #Connect-PnPOnline -Url $config.rootSiteUrl -Credentials $global:cred #-Scopes Group.ReadWrite.All
        }

        $teamsItems = $sitesData | Where { $_.EntityType -eq "MSTeam" }
        if ($null -ne $teamsItems -and $teamsItems.Length -gt 0)
        {
            # If we have any sites to provision that are MSTeam entities, connect to Teams
            Write-Log "Connect-MicrosoftTeams" -WriteToHost
            Connect-MicrosoftTeams -Credential $global:cred
        }

        # Loop through list to create sites 
        foreach ($line in $sitesData)
        {
            $siteUrl = Get-UrlByEntityType $line.EntityType $line.Site $config

            try
            {
                
                if ($null -eq $siteStatus[$siteUrl])
                {
                    Write-Log "[$siteUrl] Did not find site in history file"
                }
                else
                {
                    Write-Log "[$siteUrl] Found site in history file.  Previous status was $($siteStatus[$siteUrl])"
                }
                
                # If we have not run this site or if it previously failed during creation
				Write-Log "$siteStatus[$siteUrl] [ProcessingStatus]  Checking if site exists"
                if ($null -eq $siteStatus[$siteUrl] -or $siteStatus[$siteUrl] -lt [ProcessingStatus]::CreatingComplete)
                {

                    $skipExisting = Get-NestedMember $config "skipExisting"
                    $existingSite = $null
                    
                    if ($skipExisting)
                    {
                        # Check if Site Already Exists
                        Write-Log "[$siteUrl] Checking if site exists"
                        $existingSite = Get-PnPTenantSite -Url $siteUrl -Connection $global:tenantConn -ErrorAction Ignore
                    }
                    

                    if (-not $skipExisting -or $null -eq $existingSite -or $null -ne $siteStatus[$siteUrl])
                    { 

                        # just doing this so we have an easy list of urls after we are done
                        $global:urls += $siteUrl
                        $siteStatus[$siteUrl] = [ProcessingStatus]::Creating

                        Write-Log "[$siteUrl] Creating site" -WriteToHost
                        Write-Log "[$siteUrl] Current status: '$($siteStatus[$siteUrl])'" -WriteNewLine

                        $description = $line.Description
                        if($line.Description -eq "")
                        {
                            $description = $line.Title
                        }

                        Write-Log "Start Invoke-Expression".PadRight(50, "*")
                        Write-Log ".\create-entity.ps1"
                        Write-Log "-ConfigFile '$($ConfigFile)'"
                        Write-Log "-Site '$($line.Site)'"
                        Write-Log "-SiteTitle '$($line.Title)'"
                        Write-Log "-SiteDescription '$($description)'"
                        Write-Log "-EntityType '$($line.EntityType)'"
                        Write-Log "-Visibility '$($line.Visibility)'"
                        Write-Log "-SkipGetCredentials -BatchMode" -WriteNewLine
                        Write-Log "-AuthMode '$($AuthMode)'" -WriteNewLine
                        
                        Invoke-Expression ".\create-entity.ps1 -ConfigFile '$($ConfigFile)' -Site '$($line.Site)' -SiteTitle '$($line.Title)' -SiteDescription '$($description)' -EntityType '$($line.EntityType)' -Visibility '$($line.Visibility)' -SkipGetCredentials -BatchMode -AuthMode '$($AuthMode)' -ClientId '$($ClientId)'"

                        Write-Log "[$siteUrl] Site creation complete" -WriteToHost -WriteNewLine

                        $siteStatus[$siteUrl] = [ProcessingStatus]::CreatingComplete


                        Write-Log "End Invoke-Expression".PadRight(50, "*") -WriteNewLine

                    }
                    else 
                    {
                        Write-Log "[$siteUrl] Site exists, will not be provisioned" -WriteToHost -WriteNewLine

                        # Set status to skipped so we do not try to provision
                        $siteStatus[$siteUrl] = [ProcessingStatus]::Skipped

                        # Write to list of existing sites
                        Write-Output "$($existingSite.Url)" >> existingsites.log
                    }
                }
                else
                {
                    Write-Log "[$siteUrl]: Previously created, skipping" -WriteToHost -WriteNewLine
                }
            }
            catch
            {
                $siteStatus[$siteUrl] = [ProcessingStatus]::ErrorCreating
                Write-Log "[$siteUrl] An error occurred creating site" -Level Error
                Write-Log "[$siteUrl] $_" -Level Error
                Write-Log "[$siteUrl] $($_.ScriptStackTrace)" -Level Error -WriteNewLine
            }
        }

        # Close any open connections before moving to the next step
        #Disconnect-OpenConnections

        # Loop through list to apply provisioning to sites 
        foreach ($line in $sitesData)
        {
            $siteUrl = Get-UrlByEntityType $line.EntityType $line.Site $config

            try
            {
                if ($siteStatus[$siteUrl] -eq [ProcessingStatus]::Complete)
                {
                    Write-Log "[$siteUrl] Previously provisioned, skipping" -WriteToHost -WriteNewLine

                    $siteStatus[$siteUrl] = [ProcessingStatus]::Skipped
                    continue
                }

                # If we have not run this site or if it previously failed during creation
                if ($siteStatus[$siteUrl] -lt [ProcessingStatus]::Complete -and $siteStatus[$siteUrl] -ne [ProcessingStatus]::ErrorCreating)
                {
                    $siteStatus[$siteUrl] = [ProcessingStatus]::Provisioning

                    Write-Log "[$siteUrl] Provisioning site" -WriteToHost
                    Write-Log "[$siteUrl] Current status: '$($siteStatus[$siteUrl])'" -WriteNewLine

                    Write-Log "Start Invoke-Expression".PadRight(50, "*")
                    Write-Log ".\provision-site.ps1"
                    Write-Log "-ConfigFile '$($ConfigFile)'"
                    Write-Log "-Site '$($line.Site)'"
                    Write-Log "-SiteTitle '$($line.Title)'"
                    Write-Log "-EntityType '$($line.EntityType)'"
                    Write-Log "-SiteType '$($line.SiteType)'"
                    Write-Log "-SkipGetCredentials -BatchMode" -WriteNewLine
                    Write-Log "-AuthMode '$($AuthMode)'" -WriteNewLine

                    Invoke-Expression ".\provision-site.ps1 -ConfigFile '$($ConfigFile)' -Site '$($line.Site)' -SiteTitle '$($line.Title)' -EntityType '$($line.EntityType)' -SiteType '$($line.SiteType)' -SkipGetCredentials -BatchMode -AuthMode '$($AuthMode)' -ClientId '$($ClientId)'"

                    Write-Log "[$siteUrl] Site provisioning complete" -WriteToHost -WriteNewLine

                    Write-Log "End Invoke-Expression".PadRight(50, "*") -WriteNewLine
                    

                    $siteStatus[$siteUrl] = [ProcessingStatus]::Complete
                }
                else
                {
                    if ($siteStatus[$siteUrl] -ne [ProcessingStatus]::ErrorCreating)
                    {
                        Write-Log "[$siteUrl] Previously provisioned, skipping" -WriteToHost -WriteNewLine
                    }
                }
            }
            catch
            {
                $siteStatus[$siteUrl] = [ProcessingStatus]::ErrorProvisioning
                Write-Log "[$siteUrl] An error occurred provisioning site" -Level Error
                Write-Log "[$siteUrl] $_" -Level Error
                Write-Log "[$siteUrl] $($_.ScriptStackTrace)" -Level Error -WriteNewLine
            }
            
        }
    }

}
finally
{
    $global:urls | Out-File "provisionedUrls.txt"

    # Write status information
    $statusFolder = New-Item -Path "./status/$statusTime" -ItemType Directory
    $siteStatus.GetEnumerator() |
        Select-Object -Property @{N='SiteUrl';E={$_.Key}},
        @{N='Status';E={$_.Value}} |
            Export-Csv -NoTypeInformation -Path "./status/$statusTime/status.csv"

    Write-Log "".PadRight(50, "=")
    Write-Log "Execution Time: [$(get-date)] [$($scriptTime.Elapsed.ToString())]"
}