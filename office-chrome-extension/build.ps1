# Package office-chrome-extension into dist/extension.zip (Chrome Web Store / sideload).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dist = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $dist | Out-Null
$zip = Join-Path $dist "extension.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }

# Ship only runtime files; dev artifacts (tests, tools, package.json) stay out.
$items = @("manifest.json", "icons", "background", "content", "lib", "offscreen", "options", "popup", "viewer", "vendor")
$paths = $items | ForEach-Object { Join-Path $root $_ } | Where-Object { Test-Path $_ }
if (-not $paths) { throw "nothing to package" }

Compress-Archive -Path $paths -DestinationPath $zip -CompressionLevel Optimal
$size = (Get-Item $zip).Length
Write-Host "Wrote $zip ($([math]::Round($size/1KB)) KB)"
