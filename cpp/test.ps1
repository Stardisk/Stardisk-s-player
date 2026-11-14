param([string]$Toolchain = 'C:\Program Files\LLVM-MinGW', [string]$RateTestFile = '', [string]$MediaTestFile = '')
$ErrorActionPreference='Stop'
$server=$null
$previousRateFile=$env:STARDISK_RATE_TEST_FILE
$previousMediaFile=$env:STARDISK_MEDIA_TEST_FILE
if($RateTestFile){$env:STARDISK_RATE_TEST_FILE=(Resolve-Path -LiteralPath $RateTestFile).Path}
if($MediaTestFile){$env:STARDISK_MEDIA_TEST_FILE=(Resolve-Path -LiteralPath $MediaTestFile).Path}
Push-Location $PSScriptRoot
try {
    New-Item -ItemType Directory -Path 'build' -Force | Out-Null
    $node=(Get-Command node -ErrorAction Stop).Source
    if(Test-Path -LiteralPath .\reverse.dll){Copy-Item -LiteralPath .\reverse.dll -Destination .\build\reverse.dll -Force}
    & $node 'tests\prepare.mjs'
    if($LASTEXITCODE){throw 'Test fixture preparation failed'}
    $serverScript=Join-Path $PSScriptRoot 'tests\serve.mjs'
    $server=Start-Process -FilePath $node -ArgumentList ('"'+$serverScript+'"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
    $flags=@('--target=x86_64-w64-windows-gnu','-std=c++20','-Oz','-fno-exceptions','-fno-rtti','-fno-stack-protector','-fno-threadsafe-statics','-fno-use-cxa-atexit','-fno-builtin','-ffunction-sections','-fdata-sections','-D_WIN32_WINNT=0x0A00')
    & "$Toolchain\bin\clang++.exe" @flags '-nostdlib' '-nostartfiles' '-nodefaultlibs' 'tests\native-tests.cpp' 'src\playback.cpp' 'src\storage.cpp' 'src\native.cpp' '-Wl,--entry,testEntry,--subsystem,console,--gc-sections' '-lkernel32' '-luser32' '-lgdi32' '-lcomctl32' '-lshell32' '-lole32' '-loleaut32' '-lpropsys' '-lstrmiids' '-luuid' '-o' 'build\native-tests.exe'
    if($LASTEXITCODE){throw 'Test compilation failed'}
    & '.\build\native-tests.exe'
    $testExit=$LASTEXITCODE
    Get-Content -LiteralPath 'qa\native-tests.log' -Encoding utf8
    if($testExit){throw "Native tests failed: $testExit"}
    & "$Toolchain\bin\clang++.exe" @flags '-nostdlib' '-nostartfiles' '-nodefaultlibs' 'tests\player-tests.cpp' 'src\playback.cpp' 'src\storage.cpp' 'src\native.cpp' '-Wl,--entry,logicTestEntry,--subsystem,console,--gc-sections' '-lkernel32' '-luser32' '-lgdi32' '-lcomctl32' '-lshell32' '-lole32' '-loleaut32' '-lpropsys' '-luxtheme' '-lstrmiids' '-luuid' '-o' 'build\player-tests.exe'
    if($LASTEXITCODE){throw 'Application test compilation failed'}
    & '.\build\player-tests.exe'
    $testExit=$LASTEXITCODE
    Get-Content -LiteralPath 'qa\player-tests.log' -Encoding utf8
    if($testExit){throw "Application tests failed: $testExit"}
} finally {
    $env:STARDISK_RATE_TEST_FILE=$previousRateFile
    $env:STARDISK_MEDIA_TEST_FILE=$previousMediaFile
    if($server -and -not $server.HasExited){Stop-Process -Id $server.Id -Force;[void]$server.WaitForExit(5000)}
    Pop-Location
}
