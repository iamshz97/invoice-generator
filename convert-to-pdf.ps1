# convert-to-pdf.ps1 — Convert output/*.docx to PDF using LibreOffice (Windows / cross-platform PowerShell)
# Usage: .\convert-to-pdf.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir  = Join-Path $ScriptDir "output"
$PdfDir     = Join-Path $ScriptDir "PDFs"

Write-Host ""
Write-Host "╔══════════════════════════════════════╗"
Write-Host "║       PDF Converter                  ║"
Write-Host "╚══════════════════════════════════════╝"
Write-Host ""

# ── Locate LibreOffice ────────────────────────────────────────────────────────
$candidates = @(
    "C:\Program Files\LibreOffice\program\soffice.exe",
    "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS via pwsh
    "/usr/bin/soffice",
    "/usr/local/bin/soffice"
)

$soffice = $null
foreach ($c in $candidates) {
    if (Test-Path $c) { $soffice = $c; break }
}

if (-not $soffice) {
    $soffice = (Get-Command soffice -ErrorAction SilentlyContinue)?.Source
}

if (-not $soffice) {
    Write-Error "LibreOffice not found.`nWindows: https://www.libreoffice.org/download/download-libreoffice/`nmacOS:   brew install --cask libreoffice"
    exit 1
}

Write-Host "Using LibreOffice: $soffice"
Write-Host ""

# Create PDFs folder if it doesn't exist
New-Item -ItemType Directory -Force -Path $PdfDir | Out-Null

# ── Find .docx files in output/ ───────────────────────────────────────────────
if (-not (Test-Path $OutputDir)) {
    Write-Host "No output/ folder found. Run .\generate.ps1 first."
    exit 1
}

$docxFiles = Get-ChildItem -Path $OutputDir -Filter "*.docx" |
    Where-Object { $_.Name -notlike ".tmp_*" } |
    Sort-Object Name

if ($docxFiles.Count -eq 0) {
    Write-Host "No .docx files found in output/. Run .\generate.ps1 first."
    exit 1
}

Write-Host "Found $($docxFiles.Count) file(s) to convert:"
foreach ($f in $docxFiles) { Write-Host "  $($f.Name)" }
Write-Host ""

# ── Ask: convert all or pick specific ones? ───────────────────────────────────
Write-Host "Convert:"
Write-Host "  1) All files (default)"
Write-Host "  2) Choose specific files"
Write-Host ""
$choice = Read-Host "Choice [1/2]"
if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }

$selectedFiles = @()

if ($choice -eq "2") {
    Write-Host ""
    for ($i = 0; $i -lt $docxFiles.Count; $i++) {
        Write-Host "  $($i+1)) $($docxFiles[$i].Name)"
    }
    Write-Host ""
    $picks = (Read-Host "Enter numbers separated by spaces (e.g. 1 3)") -split '\s+'
    foreach ($pick in $picks) {
        $idx = [int]$pick - 1
        if ($idx -lt 0 -or $idx -ge $docxFiles.Count) {
            Write-Error "Error: '$pick' is out of range."
            exit 1
        }
        $selectedFiles += $docxFiles[$idx]
    }
} else {
    $selectedFiles = $docxFiles
}

# ── Convert ───────────────────────────────────────────────────────────────────
Write-Host ""
$success = 0
$failed  = 0

foreach ($docx in $selectedFiles) {
    Write-Host -NoNewline "Converting $($docx.Name) ... "
    $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($docx.Name) + ".pdf"

    $proc = Start-Process -FilePath $soffice `
        -ArgumentList "--headless", "--convert-to", "pdf", "--outdir", "`"$PdfDir`"", "`"$($docx.FullName)`"" `
        -Wait -PassThru -WindowStyle Hidden

    if ($proc.ExitCode -eq 0) {
        Write-Host "done → $pdfName"
        $success++
    } else {
        Write-Host "FAILED"
        $failed++
    }
}

Write-Host ""
Write-Host "Converted: $success file(s)"
if ($failed -gt 0) { Write-Host "Failed:    $failed file(s)" }
Write-Host "PDFs saved to: $PdfDir"
