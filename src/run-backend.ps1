<#
.SYNOPSIS
  Start the Office Editor backend (office-backend), default http://127.0.0.1:47831.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File src\run-backend.ps1
  powershell -ExecutionPolicy Bypass -File src\run-backend.ps1 -Port 47900
.NOTES
  Keep the window open while you use the extension; press Ctrl+C to stop it.
#>
param(
  [int]$Port = 0,
  [switch]$DebugBuild
)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$backendArgs = @()
if ($Port -gt 0) { $backendArgs += @('--port', "$Port") }

$cargo = Get-Command cargo -ErrorAction SilentlyContinue
$exe = Join-Path $PSScriptRoot 'target\release\office-backend.exe'

if ($cargo) {
  Write-Host 'Building and starting office-backend (the first build takes a few minutes)...'
  $profileArgs = @('--release')
  if ($DebugBuild) { $profileArgs = @() }
  & cargo run @profileArgs -p office-backend -- @backendArgs
  exit $LASTEXITCODE
} elseif (Test-Path -LiteralPath $exe) {
  Write-Host 'cargo not found - starting the previously built office-backend.exe'
  & $exe @backendArgs
  exit $LASTEXITCODE
} else {
  Write-Host ''
  Write-Host 'Rust is not installed, so office-backend cannot be built.'
  Write-Host ' 1. Install Rust from https://rustup.rs (accept the defaults)'
  Write-Host ' 2. Open a NEW terminal and run this script again.'
  Write-Host ''
  Write-Host 'The extension still works without the backend for .docx, .xlsx, .xls, .ods,'
  Write-Host '.csv, .tsv, .md and .txt files; .doc, .rtf and .odt need the backend.'
  exit 1
}
