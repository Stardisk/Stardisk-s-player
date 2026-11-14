param(
    [string]$Toolchain = 'C:\Program Files\LLVM-MinGW',
    [string]$Output = 'stardisk.exe'
)
$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
if (-not (Test-Path -LiteralPath (Join-Path $Toolchain 'bin\clang++.exe'))) {
    $local = Get-ChildItem -LiteralPath (Join-Path $project 'tools') -Directory -Filter 'llvm-mingw-*' | Select-Object -First 1
    if (-not $local) { throw 'LLVM-MinGW compiler not found. Specify -Toolchain.' }
    $Toolchain = $local.FullName
}
$bin = Join-Path $Toolchain 'bin'
$obj = Join-Path $project 'build'
New-Item -ItemType Directory -Path $obj -Force | Out-Null
Push-Location $project
try {
    & (Join-Path $bin 'llvm-windres.exe') '-Isrc' 'src\player.rc' '-O' 'coff' '-o' 'build\player.res.o'
    if ($LASTEXITCODE) { throw 'Resource compilation failed' }
    $flags = @('--target=x86_64-w64-windows-gnu', '-std=c++20', '-Oz', '-fno-exceptions', '-fno-rtti', '-fno-stack-protector', '-fno-threadsafe-statics', '-fno-use-cxa-atexit', '-fno-builtin', '-ffunction-sections', '-fdata-sections', '-Wall', '-Wextra', '-Wno-missing-field-initializers', '-Wno-unused-parameter', '-D_WIN32_WINNT=0x0A00')
    & (Join-Path $bin 'clang++.exe') @flags '-nostdlib' '-nostartfiles' '-nodefaultlibs' 'src\player.cpp' 'src\playback.cpp' 'src\storage.cpp' 'src\native.cpp' 'build\player.res.o' '-Wl,--entry,playerEntry,--subsystem,windows,--gc-sections,--dynamicbase,--nxcompat,--high-entropy-va,--strip-all,--no-insert-timestamp,-Map,build/stardisk.map' '-lkernel32' '-luser32' '-lgdi32' '-lcomctl32' '-lshell32' '-lole32' '-loleaut32' '-lpropsys' '-luxtheme' '-lstrmiids' '-luuid' '-o' $Output
    if ($LASTEXITCODE) { throw 'C++ compilation/link failed' }
    & (Join-Path $bin 'llvm-readobj.exe') '--file-headers' '--coff-imports' $Output | Set-Content -LiteralPath 'build\imports.txt'
    $imports = Get-Content -LiteralPath 'build\imports.txt' -Raw
    if ($imports -match '(?im)Name:.*(msvcrt|ucrt|vcruntime|msvcp|libc\+\+|libgcc|libwinpthread)') { throw 'Unexpected runtime dependency in executable' }
    Get-Item -LiteralPath $Output | Select-Object FullName, Length
} finally { Pop-Location }
