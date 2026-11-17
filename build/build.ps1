param([string]$Go = 'go', [string]$Output = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $Output) { $Output = Join-Path $root 'dist\DOCX_to_PDF_Converter_v1.1.0_Windows_x64.exe' }
$Output = [IO.Path]::GetFullPath($Output)
New-Item -ItemType Directory -Force (Split-Path -Parent $Output) | Out-Null
$env:GOOS='windows'; $env:GOARCH='amd64'; $env:CGO_ENABLED='0'; $env:GOTOOLCHAIN='local'
Push-Location (Join-Path $root 'src')
try {
    & $Go build -trimpath -buildvcs=false -ldflags '-s -w -H=windowsgui -X main.applicationVersion=1.1.0' -o $Output .
    if ($LASTEXITCODE -ne 0) { throw "Go build failed: $LASTEXITCODE" }
} finally { Pop-Location }
Get-FileHash -LiteralPath $Output -Algorithm SHA256
