<#
    SCRIPT OVERVIEW:
    This script updates the Defra Intranet Site Pages library with our custom columns for the approval system

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

$site = $global:sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

$ctx = Get-PnPContext

$CTName = "Site Page"
$ct = Get-PnPContentType -Identity $CTName

$list = "Site Pages"

if($null -eq $ct)
{
    throw "The content type $CTName could not be found on the site"
}

$ctx.Load($ct.FieldLinks)
$ctx.ExecuteQuery()

# Site-specific variable configuration.
$fieldNames = @("OrganisationIntranetsContentEditorInput","PageApprovalInfo")

# FIELDS - ADD FIELDS TO SITE PAGE LIBRARY
Write-Host "`nADDING FIELDS TO THE CONTENT TYPE $CTName" -ForegroundColor Green

foreach($fieldName in $fieldNames)
{
    $field = $ct.FieldLinks | Where-Object { $_.Name -eq $fieldName }

    if($null -eq $field)
    {
        Add-PnPFieldToContentType -Field $fieldName -ContentType $ct
        Write-Host "THE FIELD '$fieldName' HAS BEEN ADDED TO THE CONTENT TYPE: $CTName" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$fieldName' ALREADY EXISTS ON THE CONTENT TYPE: $CTName" -ForegroundColor Yellow 
    }
}

# Customise the "PageApprovalInfo" column for this library.
$fieldInternalName = $fieldNames[1]

$field = Get-PnPField -List $list -Identity $fieldInternalName -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{
        Hidden = $true
    }

    Write-Host "FIELD '$fieldInternalName' UPDATED" -ForegroundColor Green
}
else
{
    Write-Host "THE FIELD '$fieldInternalName' DOES NOT EXIST IN THE LIBRARY '$displayName'" -ForegroundColor Red
}

# Customise the existing "OrganisationIntranets" column for this library. The new column will be taking over user interaction.
$fieldInternalName = "OrganisationIntranets"
$newFieldInternalName = $fieldNames[0]

$field = Get-PnPField -List $list -Identity $fieldInternalName -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{
        Title = "Organisation (Intranets) - Approving ALBs"
    }

    Write-Host "FIELD '$fieldInternalName' UPDATED" -ForegroundColor Green
}
else
{
    Write-Host "THE FIELD '$fieldInternalName' DOES NOT EXIST IN THE LIBRARY '$displayName'" -ForegroundColor Red
}

# Update the views
if($null -ne $list -and $null -ne $field)
{
    # Remove the old Organisation (Intranet) field. The reason for this is SharePoint is going to manage this away from the user now
    $views = Get-PnPView -List $list | Where-Object { $_.Title -ne "" }

    foreach($view in $views)
    {
        $ctx.Load($view.ViewFields)
        $ctx.ExecuteQuery()

        $fieldExists = $view.ViewFields | Where-Object { $_ -eq $fieldInternalName }

        if($null -ne $fieldExists)
        {
            $fieldNames = New-Object Collections.Generic.List[String]

            foreach($viewField in $view.ViewFields)
            {
                if($viewField -ne $fieldInternalName)
                {
                    $fieldNames.Add($viewField)
                }
            }

            $view = Set-PnPView -List $list -Identity $view.Title -Fields $fieldNames
            Write-Host "THE FIELD '$($fieldInternalName)' HAS REMOVED FROM THE '$list' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green 
        }
        else
        {
            Write-Host "THE FIELD '$($fieldInternalName)' HAS ALREADY BEEN REMOVED FROM THE VIEW '$($view.Title)'" -ForegroundColor Yellow
        }
    }

    foreach($view in $views)
    {
        $ctx.Load($view.ViewFields)
        $ctx.ExecuteQuery()

        foreach($fieldName in $fieldNames)
        {
            $fieldExists = $view.ViewFields | Where-Object { $_ -eq $fieldName }

            if($null -eq $fieldExists)
            {
                $viewFieldNames = New-Object Collections.Generic.List[String]

                foreach($viewField in $view.ViewFields)
                {

                    $viewFieldNames.Add($viewField)
                }

                $viewFieldNames.Add($fieldName)

                $view = Set-PnPView -List $list -Identity $view.Title -Fields $viewFieldNames
                Write-Host "THE FIELD '$($fieldName)' HAS BEEN ADDED TO THE '$list' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green 
            }
            else
            {
                Write-Host "THE FIELD '$($fieldName)' HAS ALREADY BEEN ADDED TO THE VIEW '$($view.Title)'" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host ""

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript