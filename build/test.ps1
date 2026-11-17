param([string]$Go = 'go')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$work=Join-Path $root '.build'
New-Item -ItemType Directory -Force $work | Out-Null
$env:GOOS='windows';$env:GOARCH='amd64';$env:CGO_ENABLED='0';$env:GOTOOLCHAIN='local'
Push-Location (Join-Path $root 'src')
try {
    & $Go test -c -o "$work\UnitTests.exe" .
    if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed' }
    & "$work\UnitTests.exe" '-test.v=true' '-test.timeout=2m'
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed' }
    & $Go vet ./...
    if ($LASTEXITCODE -ne 0) { throw 'Go vet failed' }
} finally { Pop-Location }
