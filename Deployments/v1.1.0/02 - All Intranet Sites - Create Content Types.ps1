<#
    SCRIPT OVERVIEW:
    This script creates our custom content types for the Defra and ALB Intranet SharePoint site

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the Defra and ALB Intranet SharePoint sites
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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$false},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}
foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
    Write-Host ""

    $ctName = "Content Submission Request - Stage 2"
    $ct = Get-PnPContentType -Identity $ctName -ErrorAction SilentlyContinue

    if($null -eq $ct)
    {
        $parentCT = Get-PnPContentType -Identity Item
        $ct = Add-PnPContentType -Name $ctName -ContentTypeId "0x010047807CA071395E44BF168B9CF766B7F5" -Description "Used by 'Internal Comms Intranet Content Submissions' list to show fields that are only relevant after a submission"
        Write-Host "SITE CONTENT TYPE INSTALLED: $displayName" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE CONTENT TYPE ALREADY INSTALLED: $displayName" -ForegroundColor Yellow   
    }

    $ctFields = Get-PnPProperty -ClientObject $ct -Property Fields

    # ADD OUR CUSTOM FIELDS

    # Site-specific variable configuration.
    switch ($site.Abbreviation)
    {
        "Defra" { 
            $fieldNames = @("AltContact","ContentTypes","OrganisationIntranets","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription","AssignedTo")
        }
        default { 
            $fieldNames = @("AltContact","ContentTypes","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription","AssignedTo")
        }
    }

    foreach($field in $fieldNames)
    {
       $field = Get-PnPField $field
       $exists = $ctFields | Where-Object {$_.Id -eq $field.Id}

       if($null -eq $exists)
       {
            Add-PnPFieldToContentType -Field $field -ContentType $ct
            Write-Host "THE FIELD '$($field.Title)' HAS BEEN ADDED TO THE CONTENT TYPE '$ctName'" -ForegroundColor Green
       }
       else
       {
            Write-Host "THE FIELD '$($field.Title)' EXISTS ON THE CONTENT TYPE '$ctName' ALREADY" -ForegroundColor Yellow
       }
    }

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript