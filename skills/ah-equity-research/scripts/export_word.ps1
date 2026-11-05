<#
.SYNOPSIS
Update a generated DOCX in an independent hidden Word instance and export PDF.
.DESCRIPTION
Requires Windows, desktop Microsoft Word, KaiTi and Times New Roman. Updates
layout/reference fields, TOC and fonts, saves the supplied DOCX, then exports PDF.
Only the document and Word instance created by this script are closed. External
link/database fields are not refreshed. Final PDF rendering and visual review
are separate steps. Do not pass an original source model or an unrelated file.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DocxPath,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$PdfPath
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This exporter requires Windows and desktop Microsoft Word.' }
$docxFull = [IO.Path]::GetFullPath($DocxPath)
$pdfFull = [IO.Path]::GetFullPath($PdfPath)
if (-not (Test-Path -LiteralPath $docxFull -PathType Leaf)) { throw "DOCX not found: $docxFull" }
if ([IO.Path]::GetExtension($docxFull) -ine '.docx') { throw 'DocxPath must name a .docx file.' }
if ([IO.Path]::GetExtension($pdfFull) -ine '.pdf') { throw 'PdfPath must name a .pdf file.' }

Add-Type -AssemblyName System.Drawing
$kaitiName = ([char]0x6977).ToString() + [char]0x4F53
$installedFonts = New-Object System.Drawing.Text.InstalledFontCollection
try {
    $fontNames = @($installedFonts.Families | ForEach-Object { $_.Name })
    if (-not ($fontNames -contains 'KaiTi' -or $fontNames -contains $kaitiName)) {
        throw 'KaiTi is not installed. Install a licensed font before exporting; no substitute is used.'
    }
    if ($fontNames -notcontains 'Times New Roman') {
        throw 'Times New Roman is not installed. Install a licensed font before exporting.'
    }
} finally { $installedFonts.Dispose() }

function Update-ReportFields($Document) {
    # These fields resolve local layout/bookmarks; do not refresh external data.
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            foreach ($field in $range.Fields) {
                $code = [string]$field.Code.Text
                if ($code -match '^\s*(TOC|TOA|PAGE|NUMPAGES|PAGEREF|REF|SEQ|SECTION|SECTIONPAGES)\b') {
                    if (-not $field.Update()) { throw "Word could not update field: $code" }
                }
            }
            $range = $range.NextStoryRange
        }
    }
}

function Set-ReportFonts($Document) {
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            $range.Font.NameAscii = 'Times New Roman'
            $range.Font.NameOther = 'Times New Roman'
            $range.Font.NameBi = 'Times New Roman'
            $range.Font.NameFarEast = $kaitiName
            $range = $range.NextStoryRange
        }
    }
}

$existingWordPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$wordInstance = $null
$reportDoc = $null
$ownsWordInstance = $false
$exportWordPid = $null
$exportPageCount = $null
try {
    $wordInstance = New-Object -ComObject Word.Application
    # Word.Application has no portable Hwnd before a document window exists.
    # Require one newly created process and an empty document collection before
    # changing visibility or opening the report. Never attach to a running app.
    $newWordProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $existingWordPids -notcontains $_.Id })
    if ($newWordProcesses.Count -ne 1 -or $wordInstance.Documents.Count -ne 0) {
        throw 'Could not identify one new empty Word instance; existing user instances were left unchanged.'
    }
    $ownsWordInstance = $true
    $exportWordPid = $newWordProcesses[0].Id
    $wordInstance.Visible = $false
    $wordInstance.DisplayAlerts = 0
    $wordInstance.AutomationSecurity = 3
    $reportDoc = $wordInstance.Documents.Open($docxFull, $false, $false, $false)
    if ($reportDoc.ReadOnly) { throw 'DOCX opened read-only; close its other editing session and retry.' }

    Update-ReportFields $reportDoc
    Set-ReportFonts $reportDoc
    $reportDoc.Repaginate()
    # Font normalization can move page references after the initial TOC update.
    Update-ReportFields $reportDoc
    Set-ReportFonts $reportDoc
    $reportDoc.Repaginate()
    foreach ($toc in $reportDoc.TablesOfContents) { $toc.UpdatePageNumbers() }
    Set-ReportFonts $reportDoc
    $reportDoc.Repaginate()
    $reportDoc.Save()

    $pdfDirectory = [IO.Path]::GetDirectoryName($pdfFull)
    if (-not (Test-Path -LiteralPath $pdfDirectory -PathType Container)) {
        [void][IO.Directory]::CreateDirectory($pdfDirectory)
    }
    $reportDoc.ExportAsFixedFormat($pdfFull, 17)
    if (-not (Test-Path -LiteralPath $pdfFull -PathType Leaf) -or (Get-Item -LiteralPath $pdfFull).Length -eq 0) {
        throw 'Word returned without creating a non-empty PDF.'
    }
    $exportPageCount = $reportDoc.ComputeStatistics(2)
} finally {
    if ($null -ne $reportDoc) {
        try { $reportDoc.Close(0) } catch { Write-Warning 'Could not close this export document cleanly.' }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportDoc)
    }
    if ($null -ne $wordInstance) {
        if ($ownsWordInstance) {
            try { $wordInstance.Quit() } catch {
                if (Get-Process -Id $exportWordPid -ErrorAction SilentlyContinue) {
                    Write-Warning ('Could not quit the export Word instance: ' + $_.Exception.Message)
                }
            }
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordInstance)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Hash the DOCX after Word releases its file lock.
[pscustomobject]@{
    status = 'exported'
    docx = $docxFull
    pdf = $pdfFull
    pages = $exportPageCount
    docx_sha256 = (Get-FileHash -LiteralPath $docxFull -Algorithm SHA256).Hash.ToLowerInvariant()
    pdf_sha256 = (Get-FileHash -LiteralPath $pdfFull -Algorithm SHA256).Hash.ToLowerInvariant()
    visual_review = 'not performed by exporter'
} | ConvertTo-Json
