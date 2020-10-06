# About this Repository
This repository contains scripts and assets used to create site collection, install custom webparts, and apply PnP Provisioning Templates.

All scripts rely on the PnP PowerShell cmdlets and must be run on a Windows machine.

# Required Setup
## Install PnP
If not already installed, please install the latest PnP PowerShell cmdlets by opening a PowerShell prompt and running the following command.
```powershell
Install-Module SharePointPnPPowerShellOnline
```

If you already have it installed, you may need to update the cmdlets. 
```powershell
Update-Module SharePointPnPPowerShellOnline
```

Full instructions for PnP PowerShell can be found here: [PnP PowerShell overview (Microsoft Docs)](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)
## Config File
A valid `JSON` configuration file is required for all scripts to run properly. Each developer should create their own config file to accomodate different developer environments.

Here's an example of a valid configuration file.

```json
{
    "adminSiteUrl":  "https://contoso-admin.sharepoint.com",
    "rootSiteUrl":  "https://contoso.sharepoint.com",
    "communicationSiteDefaultPath": "sites",
    "teamSiteDefaultPath": "sites",
    "additionalSiteCollectionAdmins": "",
    "webparts": {
        "pathToFolder": "../webparts",
        "files": [
            {
                "fileName": "archived-comments.sppkg",
                "deployToTenant": true
            },
            {
                "fileName": "enhancedtext-webpart.sppkg",
                "deployToTenant": true
            }
        ]
    }
}
```

# apply-template.ps1
This script will use the PnP PowerShell Commandlets and SharePoint Online Commandlets to apply templates to an existing site collection.  It will apply a 'common' base template and then the specific one based on the provided site type.  Additionally, it will turn off the NoScript option on a modern site that is set by default so that we can apply web properties.
```powershell
./apply-template.ps1
    -configFile <String>
    -siteUrl <String>
        -skipGetCredentials <Switch>
        -batchMode <Switch>
```

## Examples
```powershell
./apply-template.ps1 -configFile './config/prod.json' -siteUrl 'https://contoso.sharepoint.com/sites/templatetest'`
```

## Parameters
    -configFile
Relative path of the configureation file to use
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:|False|
| Accept wildcard characters:|False|
<br/>

    -siteUrl
Full url for the site to apply the template to
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -skipGetCredentials
Will not prompt you for credentials. Requires a global `PSCredential` object be defined named `$global:cred`
|||
|-|-|
| Type :| Switch |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -batchMode
Prevents `Write-Host` messages from being written to the console.

|||
|-|-|
| Type :| Switch |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

# apply-template-batch.ps1
This script will loop through the sites defined in the provided site file and construct urls from the config file and use this to call the apply-template.ps1 to apply the appropriate template to the site

```powershell
./apply-template.batch.ps1
    -configFile <String>
    -sitesFile <String>
```

## Examples
```powershell
.\apply-template-batch.ps1 -configFile './config/prod.json' -sitesFile './config/test-template-sites.csv'
```

## Parameters
    -configFile
Relative path of the configureation file to use

|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -sitesFile
Relative path of a csv file containing a collection of site information

|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

# create-site.ps1
This script will create a site or team based on the parameters passed in.  If the SharePoint Url
for the site exists it will be skipped.

```powershell
./create-site.ps1
    -configFile <String>
    -site <String>
    -siteTitle <String>
    -siteDescription <String>
    -siteType <String>
        [-visibility <String>]
        [-skipGetCredentials <Switch>]
        [-batchMode <Switch>]
```

## Examples
```powershell
./create-site.ps1 -configFile './config/prod.json' -site 'contoso-templatetest' -siteTitle 'contoso Template Test' -siteDescription` 'Test Description' -siteType 'CommunicationSite'
```

## Parameters
    -configFile
Relative path of the configureation file to use
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |

    -placeID
The ID of the source Jive Place.
|||
|-|-|
| Type :| Int |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |

    -site
The site specific portion of the url (the part after /sites/ or /teams/)
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |

    -siteTitle
The display name of the site collection to be created.
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |

    -siteDescription
Description of the site collection to be created.
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -siteType
The type of site being created. Must be one of the following values: *CommunicationSite*, *TeamSite*, *MSTeam*
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -visibility
Visibility setting for newly created Microsoft Teams. Must be one fo the following values: *Public*, *Private*
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -skipGetCredentials
Will not prompt you for credentials. Requires a global `PSCredential` object be defined named `$global:cred`
|||
|-|-|
| Type :| Switch |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -batchMode
Prevents `Write-Host` messages from being written to the console.
|||
|-|-|
| Type :| Switch |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

# create-site-batch.ps1
This script will create sites and teams based on the provided csv file.  It depends on the create-site.ps1 script that is invoked when this loops through the rows in the csv and constructs the call

```powershell
    ./create-site-batch.ps1
        -configFile <String>
        -sitesFiles <String>
```

## Examples
```powershell
./create-site-batch.ps1 -configFile './config/prod.json' -sitesFile './config/test-template-sites.csv'
```

## Parameters
    -configFile
Relative path of the configureation file to use
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -sitesFile
Relative path of a csv file containing a collection of site information
|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

# install-webpart.ps1
This script can be used to deploy apps to either a Tenant or Site app catalog. It requires a valid configuration file, which has been documented in the readme. 

Deploying to the tenant app requires the SharePoint admin role and will make the webpart available for use in all sites in the tenant. 

Deploying to the site level requires the Site Admin permission and will make the webpart available only for the specified sites.

```powershell
./install-webpart.ps1 
    -configFile <String>
        [-Credentials <PSCredential>]
        [-UserName <String>]
```
<br/>

```powershell
./install-webpart.ps1 
    -configFile <String>
    -SiteUrl <String[]>]
        [-Credentials <PSCredential>]
        [-UserName <String>]
```
<br/>

## Examples
<p><strong>------------------EXAMPLE 1------------------</strong></p>
<pre><code class="lang-powershell">./create-site-batch.ps1 -configFile './config/willdev.json' -Credentials $myCredentials</code></pre>
Deploy all tenant scoped apps to the tenant using the specified credentials object.
<br/><br/>

<p><strong>------------------EXAMPLE 2------------------</strong></p>
<pre><code>./create-site-batch.ps1 -configFile './config/willdev.json' -SiteUrl 'https://contoso.sharepoint.com/sites/test-site'</code></pre>
Deploy all site scoped apps to the specified site. Will prompt user for credentials.
<br/><br/>

<p><strong>------------------EXAMPLE 3------------------</strong></p>
<pre><code>./create-site-batch.ps1 -configFile './config/willdev.json' -SiteUrl '/sites/test-site-1', '/sites/test-site-2' -Credentials $myCredentials</code></pre>
Deploy all site scoped apps to multiple sites using the specified credentials.
<br/><br/>

<p><strong>------------------EXAMPLE 4------------------</strong></p>
<pre><code>'/sites/test-site-1', '/sites/test-site-2' | ./create-site-batch.ps1 -configFile './config/willdev.json' -Credentials $myCredentials</code></pre>
Deploy all site scoped apps to multiple sites piped in using the specified credentials.
<br/>

## Parameters

    -ConfigFile

Relative path the JSON configuration file to be used.

|||
|-|-|
| Type :| String |
| Position: | Named |
| Required:| True |
| Accept pipeline input:|False|
| Accept wildcard characters:|False|
<br/>

    -SiteUrl
OPTIONAL. The URLs of sites to deploy site scoped apps to. URLs can be relative (/sites/your-site) or full (https://contoso.sharepoint.com/sites/your-site)
    
If one or more URLS are provided, only webparts not scoped to the tenant will be deployed. This parameter accepts input from the pipeline.

|||
|-|-|
| Type :| String [ ] |
| Position: | 0 |
| Required:| True |
| Accept pipeline input:|True|
| Accept wildcard characters:|False|
<br/>

    -Credentials
OPTIONAL. Can be used to pass a PSCredential that represents the M365 login credentials of the user running the script. If provided, the script will not prompt the user to login. 

|||
|-|-|
| Type: | PSCredential |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>

    -UserName
OPTIONAL. If passed, will only prompt the user for a password (unless the Credentials parameter was used).

|||
|-|-|
| Type: | String |
| Position: | Named |
| Required:| False |
| Accept pipeline input:| False |
| Accept wildcard characters:| False |
<br/>