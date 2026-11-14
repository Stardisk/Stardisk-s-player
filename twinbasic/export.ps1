param()

$ErrorActionPreference = 'Stop'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw 'Install Node.js 18 or newer, then run this script again.'
}
& node (Join-Path $PSScriptRoot 'tools/export-source.mjs')
if ($LASTEXITCODE -ne 0) { throw "Project export failed (exit $LASTEXITCODE)." }
