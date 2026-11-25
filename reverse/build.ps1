param(
    [string]$Toolchain = 'C:\Program Files\LLVM-MinGW',
    [string]$Output = 'reverse.dll',
    [switch]$OriginalMemoryProbe
)
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $Toolchain 'bin\clang++.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'LLVM-MinGW compiler not found. Specify -Toolchain.' }
$outputPath = if ([System.IO.Path]::IsPathRooted($Output)) { $Output } else { Join-Path $PSScriptRoot $Output }
New-Item -ItemType Directory -Path (Split-Path -Parent $outputPath) -Force | Out-Null
$probeFlag = if ($OriginalMemoryProbe) { '-DREVERSE_CONTINUE_MEMORY_PROBE=0' } else { '-DREVERSE_CONTINUE_MEMORY_PROBE=1' }
Push-Location $PSScriptRoot
try {
    & $compiler --target=x86_64-w64-windows-gnu -std=c++20 -Oz -Wall -Wextra -static -shared $probeFlag reverse.cpp -o $outputPath '-Wl,--strip-all,--no-insert-timestamp' -lmfreadwrite -lmfplat -lmfuuid -lole32 -lpropsys -lwinmm -luuid -lshlwapi
    if ($LASTEXITCODE) { throw 'reverse.dll compilation failed' }
    Get-Item -LiteralPath $outputPath | Select-Object FullName, Length
} finally { Pop-Location }
