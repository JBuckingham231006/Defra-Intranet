<#
    SCRIPT OVERVIEW:
    This PowerShell module is the global configuration file for this deployment.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    REQUIREMENTS:
    The global settings for all scripts can be found at the top of this script. For any environment-specific settings, please set these within the appropriate case statement of the Invoke-Configuration method.
#>

# GLOBAL SETTINGS
$global:environment = 'local001' # Which environment are we targetting?

# GLOBAL VARIABLES

# GLOBAL SCRIPT SETTINGS
if($global:PSScriptRoot.Length -gt 0)
{
    New-Item -ItemType Directory -Force -Path "$global:PSScriptRoot\Logs" | Out-Null
}
else
{
    New-Item -ItemType Directory -Force -Path "./Logs" | Out-Null
}

function Invoke-Configuration
{
    param (
        [string]$env = $global:environment
    )

    # TENANT-SPECIFC SETTINGS
    switch($env) {
        dev {
            $global:adminURL = 'https://defradev-admin.sharepoint.com'
            $global:rootURL = 'https://defradev.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = 'APHAIntranet'
                    'RelativeURL' = 'sites/APHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'ContentTypeHub'
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ''
                    'GroupPrefix' = ''
                    'RelativeURL' = 'sites/ContentTypeHub'
                    'SiteType' = 'System'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'DefraIntranet '
                    'RelativeURL' = 'sites/defraintranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'EAIntranet'
                    'RelativeURL' = 'sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = 'MMOIntranet'
                    'RelativeURL' = 'sites/MMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = 'NEIntranet'
                    'RelativeURL' = 'sites/NEIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }
      
        preprod {
            $global:adminURL = 'https://defra-admin.sharepoint.com'
            $global:rootURL = 'https://defra.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/PPAPHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'ContentTypeHub'
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ''
                    'GroupPrefix' = ''
                    'RelativeURL' = 'sites/ContentTypeHub'
                    'SiteType' = 'System'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/PPDefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/PPEAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/PPMMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/PPNEIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }

        production {
            $global:adminURL = 'https://defra-admin.sharepoint.com'
            $global:rootURL = 'https://defra.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/APHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'ContentTypeHub'
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ''
                    'GroupPrefix' = ''
                    'RelativeURL' = 'sites/ContentTypeHub'
                    'SiteType' = 'System'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/DefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/MMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = ''
                    'RelativeURL' = '/sites/NEIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }

        # JB Development Environment
        local001 {
            $global:adminURL = 'https://buckinghamdevelopment-admin.sharepoint.com'
            $global:rootURL = 'https://buckinghamdevelopment.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = ''
                    'RelativeURL' = ''
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'ContentTypeHub'
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ''
                    'GroupPrefix' = ''
                    'RelativeURL' = 'sites/ContentTypeHub'
                    'SiteType' = 'System'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'HUBSITE001'
                    'RelativeURL' = 'sites/DefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'DEFRA001'
                    'RelativeURL' = 'sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = ''
                    'RelativeURL' = ''
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = ''
                    'RelativeURL' = ''
                    'SiteType' = 'ALB'
                }
            )
        }

        # NM Development Environment
        local002 {
            $global:adminURL = 'https://nhmsolutions-admin.sharepoint.com'
            $global:rootURL = 'https://nhmsolutions.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = 'APHA'
                    'RelativeURL' = 'sites/apha'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'ContentTypeHub'
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ''
                    'GroupPrefix' = 'APHA Intranet'
                    'RelativeURL' = 'sites/ContentTypeHub'
                    'SiteType' = 'System'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'Defra Intranet'
                    'RelativeURL' = 'sites/defraintranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'EA Intranet'
                    'RelativeURL' = 'sites/eaintranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = 'MMO Intranet'
                    'RelativeURL' = 'sites/mmointranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = 'NE Intranet'
                    'RelativeURL' = 'sites/neintranet'
                    'SiteType' = 'ALB'
                }
            )
        }
    }
}