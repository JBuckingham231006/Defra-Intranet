<#
    SCRIPT OVERVIEW:
    This script creates our custom submission list

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the Defra Intranet SharePoint site
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

$site = $global:sites | Where-Object { $_.Abbreviation -eq "Defra" -and $_.RelativeURL.Length -gt 0 }

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
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
$fieldNames = @("AltContact","ContentTypes","OrganisationIntranets","LineManager","PublishBy","StakeholdersInformed","ContentSubmissionStatus","ContentSubmissionDescription","AssignedTo")

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

# Customise the OrganisationIntranets column for this list
$field = Get-PnPField -List $list -Identity "OrganisationIntranets" -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{
        Title = "Content Relevant To"
        Description = "Select whether the content is relevant to the whole of the Defra group or specific departments or functions"
        Required = $true
    }
}
else
{
    Write-Host "THE FIELD 'OrganisationIntranets' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
}

# Customise the ContentSubmissionStatus column for this list
$field = Get-PnPField -List $list -Identity "ContentSubmissionStatus" -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{
        Hidden = $true
        CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Pending Approval"]},"sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont sp-field-fontSizeSmall sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved"]},"sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Rejected"]},"sp-css-backgroundColor-BgDustRose sp-css-borderColor-DustRoseFont sp-field-fontSizeSmall sp-css-color-DustRoseFont",""]}]}]}},"txtContent":"[$ContentSubmissionStatus]"}]}'
    }
}
else
{
    Write-Host "THE FIELD 'ContentSubmissionStatus' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
}

# Customise the AssignedTo column for this list
$field = Get-PnPField -List $list -Identity "AssignedTo" -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{"SelectionMode"=0}
}
else
{
    Write-Host "THE FIELD 'AssignedTo' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
}

# Update the list's default view with our new fields
$view = Get-PnPView -List $list -Identity "All Items"

if($null -ne $view)
{
    $fieldNames = @("Attachments","LinkTitle","AssignedTo","ContentSubmissionDescription","Author","ContentSubmissionStatus","PublishBy","ContentTypes","AltContact","OrganisationIntranets","LineManager","StakeholdersInformed")
    $view = Set-PnPView -List $list -Identity $view.Title -Fields $fieldNames
    Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
}

# Set unique permissions for the list so anyone on the site can add an item
if($null -ne $site.GroupPrefix -and $site.GroupPrefix.Length -gt 0)
{
    Write-Host "`nCUSTOMISING LIST PERMISSIONS" -ForegroundColor Green
    Set-PnpList -Identity $list -BreakRoleInheritance
    
    Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Owners" -AddRole "Full Control"
    Write-Host "'$($site.GroupPrefix) Owners' given Full Control" -ForegroundColor Yellow
    
    Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Members" -AddRole "Edit"
    Write-Host "'$($site.GroupPrefix) Members' given Edit permissions to the list" -ForegroundColor Yellow

    Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Visitors" -AddRole "Contribute"
    Write-Host "'$($site.GroupPrefix) Visitors' given Contribute permissions to the list" -ForegroundColor Yellow
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript