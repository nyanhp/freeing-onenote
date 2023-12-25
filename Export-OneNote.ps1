
<#PSScriptInfo

.VERSION 1.1.0

.GUID b9742b5d-5b71-4a08-bbfe-635827e34076

.AUTHOR Jan-Hendrik Peters

.PSEDITION Core, Desktop

.COMPANYNAME Shiftavenue GmbH

.COPYRIGHT Jan-Hendrik Peters, 2023

.TAGS OneNote, Markdown, Graph

.LICENSEURI https://raw.githubusercontent.com/nyanhp/freeing-onenote/main/LICENSE

.PROJECTURI https://github.com/nyanhp/freeing-onenote

.ICONURI

.EXTERNALMODULEDEPENDENCIES  MiniGraph, MarkdownPrince

.REQUIREDMODULES MiniGraph, MarkdownPrince

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#
.SYNOPSIS
 This script exports one or more OneNote notebooks to markdown

.DESCRIPTION
 This script exports one or more OneNote notebooks to markdown
 using the Graph API and an Entra App Registration.

 The script requires the following modules to be installed:
    - MiniGraph
    - MarkdownPrince

.PARAMETER TenantId
    The tenant id to use for the Graph API. Defaults to 'Common'.

.PARAMETER OneNoteAppClientId
    The client id of the app registration to use for the Graph API. Defaults to '812899b7-584c-4812-8aee-11d3e164d58b', which is managed by the author.

.PARAMETER User
    The user to use for the Graph API. Defaults to 'me'. Use string like user/GUID to request notebooks from other users if the permissions are set correctly.

.PARAMETER Notebook
    The name of the notebook to export. Can be specified multiple times. If not specified, all notebooks will be exported.

.PARAMETER All
    If specified, all notebooks will be exported.

.PARAMETER Path
    The path to export the notebooks to. If the path does not exist, it will be created.

.EXAMPLE
    ./Export-OneNote.ps1 -Notebook 'My Notebook' -Path "$home/freedomfromproprietaryformats"

#>

[CmdletBinding(DefaultParameterSetName = 'All')]
param
(
    [Parameter(ParameterSetName = 'Notebook')]
    [Parameter(ParameterSetName = 'All')]
    [string]
    $TenantId = 'Common',

    # Use mine or create your own
    [Parameter(ParameterSetName = 'Notebook')]
    [Parameter(ParameterSetName = 'All')]
    [string]
    $OneNoteAppClientId = '812899b7-584c-4812-8aee-11d3e164d58b',

    [string]
    $User = 'me',

    [Parameter(Mandatory = $true, ParameterSetName = 'Notebook')]
    [string[]]
    $Notebook,

    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    [switch]
    $All,

    [Parameter(Mandatory = $true, ParameterSetName = 'Notebook')]
    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    $Path
)

#requires -Module MiniGraph
#requires -Module MarkdownPrince

Connect-GraphDeviceCode -TenantId $TenantId -ClientId $OneNoteAppClientId
Set-GraphEndpoint -Type beta

$notebooks = if ($All.IsPresent)
{
    Invoke-GraphRequest -Query "$($User)/onenote/notebooks"
}
else
{
    $Notebook | ForEach-Object {
        Invoke-GraphRequest -Query "$($User)/onenote/notebooks?`$filter=displayName eq '$($_ -replace "'", "''")'"
    }
}

if (-not (Test-Path $Path))
{
    $null = New-Item -Path $Path -ItemType Directory -Force
}

$mg = Get-Module -Name MiniGraph
$token = & $mg { $script:token }

foreach ($book in $notebooks)
{
    $bookPath = Join-Path -Path $Path -ChildPath $book.displayName
    $sections = Invoke-GraphRequest -Query "$($User)/onenote/notebooks/$($book.id)/sections"
    if (-not (Test-Path -Path $bookPath))
    {
        $null = New-Item -Path $bookPath -ItemType Directory -Force
    }

    foreach ($section in $sections)
    {
        $sectionPath = Join-Path -Path $bookPath -ChildPath $section.displayName
        if (-not (Test-Path -Path $sectionPath))
        {
            $null = New-Item -Path $sectionPath -ItemType Directory -Force
        }
        $pages = Invoke-GraphRequest -Query "$($User)/onenote/sections/$($section.id)/pages"

        foreach ($page in $pages)
        {
            $pagePath = Join-Path -Path $sectionPath -ChildPath "$($page.title).md"
            $content = Invoke-GraphRequest -Query "$($User)/onenote/pages/$($page.id)/content"
            $imgCount = 0
            foreach ($image in $content.SelectNodes("//img"))
            {
                $header = @{
                    Authorization = "Bearer $token"
                }
                $imgName = '{0}_{1:d10}.png' -f $page.title, $imgCount
                $imgPath = Join-Path -Path $sectionPath -ChildPath resources
                if (-not (Test-Path -Path $imgPath))
                {
                    $null = New-Item -Path $imgPath -ItemType Directory -Force
                }

                Invoke-RestMethod -Method Get -Uri $image.'data-fullres-src' -Headers $header -OutFile (Join-Path $imgPath $imgName)

                $image.src = [uri]::EscapeUriString(('./resources/{1}' -f $section.displayName, $imgName))
            }

            $content.OuterXml | ConvertFrom-HTMLToMarkdown -DestinationPath $pagePath -Format
        }
    }
}
