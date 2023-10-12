<#
    SCRIPT OVERVIEW:
    This script creates the site columns required by our custom list(s) and libraries within the DEFRA SharePoint site

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

# "Alternative Contact" column
$displayName = "Alternative Contact"
$field = Get-PnPField -Identity "AltContact" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "User" -InternalName "AltContact" -DisplayName $displayName -Required
    Set-PnPField -Identity $field.Id -Values @{SelectionMode=0}

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Content Relevant To" column
$displayName = "Content Relevant To"
$field = Get-PnPField -Identity "ContentRelevantTo" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "MultiChoice" -InternalName "ContentRelevantTo" -DisplayName $displayName -Required -Choices "DEFRA","EA","NE","APHA","RPA","MMO","CF","VMD","Cefas","JNCC","Kew"
    Set-PnPField -Identity $field.Id -Values @{
        Description = "Select whether the content is relevant to the whole of the DEFRA group or specific departments or functions"
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Content Type" column
$displayName = "Content Types"
$field = Get-PnPField -Identity "ContentTypes" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "Choice" -InternalName "ContentTypes" -DisplayName $displayName -Required -Choices "News Story","Blog or Online Diary","Office Notice","Site Page"
    Set-PnPField -Identity $field.Id -Values @{
        Description = "Please select what kind of content you are submitting:"; 
        CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]",""]},"",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","News story"]},"sp-css-backgroundColor-successBackground50 sp-css-color-green",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Blog or online diary"]},"sp-css-backgroundColor-warningBackground50 sp-css-color-neutralPrimaryAlt",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Office notice"]},"sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary",{"operator":":","operands":[{"operator":"==","operands":["[$ContentTypes]","Site Pages"]},"sp-css-backgroundColor-BgCyan sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-CyanFont sp-css-color-CyanFont",""]}]}]}]}]}},"txtContent":"[$ContentTypes]"}],"templateId":"BgColorChoicePill"}'
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Line Manager" column
$displayName = "Line Manager"
$field = Get-PnPField -Identity "LineManager" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "User" -InternalName "LineManager" -DisplayName $displayName -Required
    Set-PnPField -Identity $field.Id -Values @{
        SelectionMode = 0;
        Description = "Please tell us who in your senior management provided the final sign-off on this story and confirm all relevant stakeholders have been informed:"
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Preferred Timing" column
$displayName = "Preferred Timing"
$field = Get-PnPField -Identity "PublishBy" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "DateTime" -InternalName "PublishBy" -DisplayName $displayName -Required
    Set-PnPField -Identity $field.Id -Values @{
        FriendlyDisplayFormat = [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType]::Relative;
        Description="Please let us know when you want your content published and any reason for that date e.g., a policy launch or awareness day. We aim to publish on the requested dates, but this may not be possible if it is short notice or there are competing internal announcements."
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Stakeholders Informed" column
$displayName = "Stakeholders Informed"
$field = Get-PnPField -Identity "StakeholdersInformed" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "Choice" -InternalName "StakeholdersInformed" -DisplayName $displayName -Required -Choices "Yes","No"
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Status" column
$displayName = "Status"
$field = Get-PnPField -Identity "ContentSubmissionStatus" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "Choice" -InternalName "ContentSubmissionStatus" -DisplayName $displayName -Choices "Pending Approval","Approved","Rejected"
    Set-PnPField -Identity $field.Id -Values @{
        DefaultValue ="Pending Approval"
        Description = "Status of the content."; 
        CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$Status]","Pending"]},"sp-css-backgroundColor-BgLightGray sp-css-borderColor-LightGrayFont sp-css-color-LightGrayFont",{"operator":":","operands":[{"operator":"==","operands":["[$Status]","Approved"]},"sp-css-backgroundColor-BgGreen sp-css-borderColor-WhiteFont sp-css-color-WhiteFont",{"operator":":","operands":[{"operator":"==","operands":["[$Status]","Rejected"]},"sp-css-backgroundColor-BgPeach sp-css-borderColor-PeachFont sp-css-color-PeachFont",{"operator":":","operands":[{"operator":"==","operands":["[$Status]",""]},"","sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary"]}]}]}]}},"txtContent":"[$Status]"}],"templateId":"BgColorChoicePill"}'
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Text" column
$displayName = "Text"
$field = Get-PnPField -Identity "ContentSubmissionDescription" -ErrorAction SilentlyContinue
if($null -eq $field)
{
    $field = Add-PnPField -Type "Note" -InternalName "ContentSubmissionDescription" -DisplayName $displayName -Required
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript