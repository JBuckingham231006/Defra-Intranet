<#
    SCRIPT OVERVIEW:
    This script creates our custom submission list

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the DEFRA Intranet SharePoint site
    OR
    - Access to the SharePoint Tenant Administration site
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

$site = $global:sites | Where-Object { $_.Abbreviation -eq "DEFRA" -and $_.RelativeURL.Length -gt 0 }

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'DEFRA Intranet' or is not configured correctly"
}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

# Create new "Submission" list
$displayName = "Internal Comms Intranet Content Submissions"
$listURL = "Lists/ICICS"

$list = Get-PnPList -Identity $listURL


if($null -eq $list)
{
    $list = New-PnPList -Template GenericList -Title $displayName -Url $listURL -EnableVersioning
    Write-Host "LIST CREATED: $displayName (URL: $listURL)" -ForegroundColor Green
}
else
{
    Write-Host "THE LIST '$displayName' ALREADY EXISTS" -ForegroundColor Yellow
}

# Add our custom columns to the list
$fieldNames = @("AltContact","ContentTypes","ContentRelevantTo","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription")

foreach($fieldName in $fieldNames)
{
    $field = Get-PnPField -List $list -Identity $fieldName -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        Add-PnPField -List $list -Field $fieldName
        Write-Host "FIELD ADDED TO THE '$displayName' LIST: $fieldName" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$fieldName' ALREADY EXISTS IN THE LIST '$displayName'" -ForegroundColor Yellow 
    }
}

# Update the list's default view with our new fields
$view = Get-PnPView -List $list -Identity "All Items"

if($null -ne $view)
{
    $fieldNames = @("LinkTitle","ContentSubmissionDescription","Author","ContentSubmissionStatus","PublishBy","ContentTypes","Attachments","AltContact","ContentRelevantTo","LineManager","StakeholdersInformed")
    $view = Set-PnPView -List $list -Identity $view.Title -Fields $fieldNames
    Write-Host "LIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
