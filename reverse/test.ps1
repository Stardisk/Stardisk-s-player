param([string]$Toolchain = 'C:\Program Files\LLVM-MinGW')
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1') -Toolchain $Toolchain
Push-Location $PSScriptRoot
try {
    New-Item -ItemType Directory -Path build,build/no-dll,qa -Force | Out-Null
    # Keep the script ASCII-compatible with Windows PowerShell 5.1 while testing a Unicode path.
    $fixture = Join-Path $PSScriptRoot ('qa\01 - ' + [char]0x00C4 + [char]0x044F + '.wav')
    # Thirty seconds of silent mono PCM16 at 44100 Hz; no user audio is used.
    $dataSize = 30 * 44100 * 2
    $writer = [System.IO.BinaryWriter]::new([System.IO.File]::Create($fixture))
    try {
        $writer.Write([System.Text.Encoding]::ASCII.GetBytes('RIFF'))
        $writer.Write([int](36 + $dataSize))
        $writer.Write([System.Text.Encoding]::ASCII.GetBytes('WAVEfmt '))
        $writer.Write([int]16)
        $writer.Write([short]1)
        $writer.Write([short]1)
        $writer.Write([int]44100)
        $writer.Write([int]88200)
        $writer.Write([short]2)
        $writer.Write([short]16)
        $writer.Write([System.Text.Encoding]::ASCII.GetBytes('data'))
        $writer.Write([int]$dataSize)
        $writer.Write([byte[]]::new($dataSize))
    } finally { $writer.Dispose() }
    & (Join-Path $Toolchain 'bin\clang++.exe') --target=x86_64-w64-windows-gnu -std=c++20 -Oz -Wall -Wextra -static -municode tests.cpp -o build/reverse-tests.exe
    if ($LASTEXITCODE) { throw 'Reverse API test compilation failed' }
    Copy-Item -LiteralPath reverse.dll -Destination build/reverse.dll -Force
    & .\build\reverse-tests.exe $fixture
    if ($LASTEXITCODE) { throw 'Reverse API tests failed' }
    Copy-Item -LiteralPath build/reverse-tests.exe -Destination build/no-dll/reverse-tests.exe -Force
    if (Test-Path -LiteralPath build/no-dll/reverse.dll) { throw 'The missing-DLL test folder contains a DLL.' }
    & .\build\no-dll\reverse-tests.exe --missing
    if ($LASTEXITCODE) { throw 'Missing-DLL test failed' }
} finally { Pop-Location }
