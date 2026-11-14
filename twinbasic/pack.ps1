param([switch]$Force)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'StardiskPlayer2NoWMP.twinproj'
$tool = Join-Path $PSScriptRoot 'tools/impexp.mjs'
$source = Join-Path $PSScriptRoot 'source'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw 'Install Node.js 18 or newer, then run this script again.'
}
$toolArgs = @($tool, 'import', $project, $source)
if ($Force) { $toolArgs += '--overwrite' }
& node @toolArgs
if ($LASTEXITCODE -ne 0) { throw "Project packing failed (exit $LASTEXITCODE)." }
Write-Host "Open in twinBASIC: $project"
