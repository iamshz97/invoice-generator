# generate.ps1 — Interactive invoice generator (Windows / cross-platform PowerShell)
# Usage: .\generate.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host ""
Write-Host "╔══════════════════════════════════════╗"
Write-Host "║        Invoice Generator             ║"
Write-Host "╚══════════════════════════════════════╝"
Write-Host ""

# ── Ask: current month or specific months? ───────────────────────────────────
Write-Host "Generate invoice for:"
Write-Host "  1) Current month (default)"
Write-Host "  2) Specific month(s)"
Write-Host ""
$choice = Read-Host "Choice [1/2]"
if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }

$monthsArg = @()

if ($choice -eq "2") {
    Write-Host ""
    Write-Host "Enter one or more months in YYYYMM format, separated by spaces."
    Write-Host "Example: 202603 202604 202605"
    Write-Host ""
    $rawInput = Read-Host "Month(s)"

    foreach ($token in $rawInput -split '\s+') {
        if ($token -notmatch '^\d{6}$') {
            Write-Error "Error: '$token' is not a valid YYYYMM value. Exiting."
            exit 1
        }
        $monthsArg += $token
    }
}

# ── Confirm before running ────────────────────────────────────────────────────
Write-Host ""
if ($monthsArg.Count -eq 0) {
    $currentMonth = Get-Date -Format "yyyyMM"
    Write-Host "Will generate invoice for: $currentMonth (current month)"
} else {
    Write-Host "Will generate invoices for: $($monthsArg -join ' ')"
}

Write-Host ""
$confirm = Read-Host "Proceed? [Y/n]"
if ([string]::IsNullOrWhiteSpace($confirm)) { $confirm = "Y" }
if ($confirm -notmatch '^[Yy]$') {
    Write-Host "Aborted."
    exit 0
}

# ── Run the .NET app ──────────────────────────────────────────────────────────
Write-Host ""
if ($monthsArg.Count -eq 0) {
    dotnet run
} else {
    dotnet run -- $monthsArg
}

Write-Host ""
Write-Host "Done. Check the output/ folder for generated files."
Write-Host "Run .\convert-to-pdf.ps1 to convert .docx files to PDF."
