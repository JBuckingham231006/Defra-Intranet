<#
    SCRIPT OVERVIEW:
    This script creates our custom column(s) in each of the Site Page libraries in the Intranet sites

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Full Control (minimum requirements) 
#>

$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

Import-Module SharePointPnPPowerShellOnline

if ($null -ne $psISE)
{
    $logfileName = $($psISE.CurrentFile.FullPath.Split('\'))[$psISE.CurrentFile.FullPath.Split('\').Count-1]
    $logfileName = $logfileName.Replace(".ps1",".txt")

    $global:scriptPath = Split-Path -Path $psISE.CurrentFile.FullPath

    Import-Module "$global:scriptPath\PSModules\Configuration.psm1" -Force
    Import-Module "$global:scriptPath\PSModules\Helper.psm1" -Force
}
else
{
    Clear-Host

    $logFileName = $MyInvocation.MyCommand.Name
    $global:scriptPath = "."

    Import-Module "./PSModules/Configuration.psm1" -Force
    Import-Module "./PSModules/Helper.psm1" -Force
}

$logfileName = $logfileName.Replace(".ps1",".txt")
Start-Transcript -path "$global:scriptPath/Logs/$logfileName" -append | Out-Null

Invoke-Configuration

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"
    Connect-PnPOnline -Url $fullURL -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $fullURL" -ForegroundColor Cyan

    if($site.Abbreviation -ne "EA")
    {
        # "Content Owner - Team" column
        $displayName = "Content Owner - Team"
        $internalName = "Content_x0020_Owner_x0020__x002d__x0020_Team"

        $field = Get-PnPField -Identity $internalName -ErrorAction SilentlyContinue

        if($null -eq $field)
        {
            $field = Add-PnPField -Type "Choice" -InternalName $internalName -DisplayName $displayName -Required
            Set-PnPField -Identity $field.Id -Values @{SelectionMode=0}

            Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
        }
        else 
        {
            Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
        }
    }
    else
    {
        Write-Host "SKIPPING SITE. WE DO NOT INSTALL THIS COLUMN ON THE '$($site.Abbreviation)' SITE" -ForegroundColor Yellow
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript