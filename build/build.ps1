# Build an independent Windows x64 .exe from the reconstructed sources.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location (Join-Path $root 'src')
try {
    $env:GOOS='windows'
    $env:GOARCH='amd64'
    $env:CGO_ENABLED='0'
    go build -trimpath -ldflags '-s -w' -o (Join-Path $root 'DOCX_to_PDF_REBUILT_x64.exe') .
    if ($LASTEXITCODE -ne 0) {
        throw "go build failed with exit code $LASTEXITCODE."
    }
} finally { Pop-Location }
Write-Host 'Build complete.'
