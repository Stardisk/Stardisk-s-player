param(
    [string]$Toolchain = 'C:\Program Files\LLVM-MinGW'
)
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $Toolchain 'bin\clang++.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'LLVM-MinGW compiler not found. Specify -Toolchain.' }
Push-Location $PSScriptRoot
try {
    & $compiler --target=x86_64-w64-windows-gnu -std=c++20 -Oz -Wall -Wextra -static -shared reverse.cpp -o reverse.dll '-Wl,--strip-all,--no-insert-timestamp' -lmfreadwrite -lmfplat -lmfuuid -lole32 -lpropsys -lwinmm -luuid -lshlwapi
    if ($LASTEXITCODE) { throw 'reverse.dll compilation failed' }
    Get-Item -LiteralPath reverse.dll | Select-Object FullName, Length
} finally { Pop-Location }
