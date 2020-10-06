[CmdletBinding(DefaultParameterSetName="PROVISION")]
PARAM(
    [Parameter(Mandatory=$true)]
    [string]$ConfigFile,
    [Parameter(Mandatory=$true)]
    [string]$DataFile = "data/sites.csv"    
)
BEGIN{    
    $sites = Import-CSV (Join-Path $script:PSScriptRoot $sitesFile)
}
PROCESS{    
    foreach($site in $sites){
        Invoke-Expression ".\Add-NewSite.ps1 -configFile '$($configFile)' -placeID '$($line.PlaceID)' -site '$($line.Site)' -siteTitle '$($line.Title)' -siteDescription '$($description)' -siteType '$($line.SiteType)' -visibility '$($line.Visibility)' -skipGetCredentials -batchMode"
    }
}
END{

}
