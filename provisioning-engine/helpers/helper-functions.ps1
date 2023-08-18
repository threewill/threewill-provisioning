# Helper function similar to lodash "get"
# Allows passing an PSObject and a dot delimited property path
# If the property is found, the value is returned, otherwise null 
function Get-NestedMember {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$InputValue,
        [Parameter(Mandatory=$true)]
        [string]$PropertyString
    )

    $splitProps = $PropertyString.Split('.')

    if ($splitProps.length -gt 1) 
    {
        $propKey = $splitProps[0]

        if ($null -ne ($InputValue | Get-Member $propKey))
        {
            $prop = $InputValue.$propKey

            return (Get-NestedMember -InputValue $prop -PropertyString ($splitProps[1..($splitProps.length - 1)] -Join '.'))
        }
        else
        {
            return $null
        }
    }
    else 
    {
        if ($null -ne ($InputValue | Get-Member $PropertyString))
        {
            return $InputValue.$PropertyString
        }
        else
        {
            return $null
        }
    }
}

# Helper function to get the full site url based on the entity type
# and settings in the config file
function Get-UrlByEntityType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EntityType,
        [Parameter(Mandatory = $true)]
        [string]$SitePath,
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        [Parameter(Mandatory = $false)]
        [string]$DefaultManagedPath = "sites"
    )

    if ($EntityType -eq "CommunicationSite")
    {
        $managedPath = Get-NestedMember $Config "communicationSiteDefaultPath"

        if ($null -eq $managedPath)
        {
            $managedPath = $DefaultManagedPath
        }

        return "$($Config.rootSiteUrl)/$($managedPath)/$($SitePath)"
    }
    elseif ($EntityType -eq "IntranetSpokeSite")
    {
        $managedPath = Get-NestedMember $Config "intranetSpokeSiteDefaultPath"

        if ($null -eq $managedPath)
        {
            $managedPath = $DefaultManagedPath
        }

        return "$($Config.rootSiteUrl)/$($managedPath)/$($SitePath)"
    }
    else
    {
        $managedPath = Get-NestedMember $config "teamSiteDefaultPath"

        if ($null -eq $managedPath)
        {
            $managedPath = $DefaultManagedPath
        }

        return "$($config.rootSiteUrl)/$($managedPath)/$($SitePath)"
    }
}

# The Disconnect-PnPOnline cmdlet throws an error if no connections are open
# We want to be able to ensure there are no open connections, therefore we 
# are using this function to ignore the error
function Disconnect-OpenConnections
{
    try
    {
        #Disconnect-PnPOnline
    }
    catch
    {
        # A connection was not open, return
        return
    }
}

function Write-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [string]$Path=$global:logFile,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info","Debug")]
        [string]$Level="Info",
        [Parameter(Mandatory=$false)]
        [switch]$WriteToHost,
        [Parameter(Mandatory=$false)]
        [switch]$WriteNewLine
    )

    # Create the file if it doesn't exist
    if(-not (Test-Path -Path $Path))
    {
        Write-Verbose "Creating '$Path'."
        $logFile = New-Item $Path -ItemType File
    }

    $formattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    switch ($Level)
    {
        "Error" {
            # This script can be used to log an Error but to actually throw an Exception 
            # call Write-Error
            # Write-Error $Warning
            Write-Host "ERROR: $Message" -ForegroundColor red
        }
        "Warn" {
            Write-Warning $Message
        }
        "Debug" {
            Write-Debug $Message
        }
        Default {
            Write-Verbose $Message
        }
    }

    # Write to log file
    "$formattedDate $($Level.ToUpper()): $Message" | Out-File -FilePath $Path -Append

    if ($WriteNewLine)
    {
        "$formattedDate $($Level.ToUpper()):" | Out-File -FilePath $Path -Append
    }

    if ($WriteToHost -and $Level -ne "Error")
    {
        Write-Host "$($Level.ToUpper()): $Message"
        if ($WriteNewLine)
        {
            Write-Host ""
        }
    }
}