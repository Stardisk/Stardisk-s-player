# Stardisk's Player — C++

**English** | [Русский](README.ru.md)

This is the main version of Stardisk's Player and the source for future releases. It is the twinBASIC version rewritten in modern C++20 with the help of AI, using the Windows API. The original VB6 versions were written without vibe coding; the twinBASIC version was developed partly with its help.

The player keeps the original idea: play files from the selected folder without a music library. It supports Unicode paths, compact mode, file operations, HTTP/HTTPS streams, multimedia keys, shuffle, track and A–B repeat, and playback speed control. In the player window, the left and right arrow keys seek by 10 seconds; time input fields keep their normal cursor behavior.

## Requirements

Windows 10/11 x64. Normal playback needs only the built EXE; no Visual C++ Redistributable, .NET or Node.js is required to run it. Playback uses system DirectShow and Media Foundation, so available formats depend on Windows codecs. The interface is in Russian.

The player writes `presets.ini` and `repeatlist.ini` beside the EXE, so that folder must be writable. Reverse playback is optional. Its DLL is not included at this stage; without it, the reverse menu item is hidden.

## Source layout

- `src/` — C++ sources, the Windows resource script and the application icon.
- `src/ui.generated.h` — checked-in control and menu definitions originally generated from the twinBASIC forms. No generation step is needed to build.
- `src/reverse_api.h` — the interface and loader for the optional DLL; the DLL implementation is not included.
- `tests/` and `test.ps1` — playback and application tests, plus scripts that create their silent audio fixtures and local HTTP server.
- `build.ps1` — builds the x64 EXE and checks its runtime imports.
- `tools/toolchain.json` — records the LLVM-MinGW distribution used during development. The compiler itself is not stored in Git.

Build output, test audio, logs, personal settings and reference exports are excluded from Git. No prebuilt EXE or DLL is included.

## Build

Install LLVM-MinGW with its x64 Windows toolchain. From the repository root:

```powershell
.\cpp\build.ps1
```

The default compiler location is `C:\Program Files\LLVM-MinGW`. To use another installation:

```powershell
.\cpp\build.ps1 -Toolchain 'C:\Tools\llvm-mingw'
```

The result is `cpp/stardisk.exe`. Intermediate files, the linker map and the import report go into `cpp/build/`. To build under another name while a player instance is running:

```powershell
.\cpp\build.ps1 -Output 'build\stardisk-check.exe'
```

The build uses `-nostdlib -nostartfiles -nodefaultlibs`. It links Windows import libraries and rejects detected CRT or third-party runtime dependencies. Node.js and twinBASIC are not needed to build the player.

## Tests

Tests additionally require Node.js in `PATH`. Run:

```powershell
.\cpp\test.ps1
```

Use `-Toolchain` if LLVM-MinGW is installed elsewhere. Tests create their own silent WAV files and settings under `cpp/qa/`, run an HTTP server on `127.0.0.1:52751`, and stop the server afterward. Test binaries go into `cpp/build/`. User music and settings are not used.

The suite checks playback, volume, speed, seeking, pause, end of track, released file handles, HTTP playback, Unicode paths, INI storage and application behavior. Without the optional DLL, it also checks that reverse controls stay unavailable.

Optional extra checks can use a local audio file opened read-only:

```powershell
.\cpp\test.ps1 -RateTestFile 'F:\Music\test.mp3'
.\cpp\test.ps1 -MediaTestFile 'F:\Music\test.mp3'
```

The rate fixture should be at least 15 seconds long. The media fixture should exercise the local Media Foundation fallback. Some system media tests require a normal Windows session rather than a restricted sandbox.

## Releases

Future player releases are built from `cpp/` on `main` at the chosen release commit or tag. Built executables belong in GitHub Releases, not in the source tree. The existing historical VB6 release is preserved.

## Verification of this import

The source-only copy builds with LLVM-MinGW. Both the native playback and application test suites passed in a normal Windows session without `reverse.dll`. Tests in the restricted sandbox could not use DirectShow and local HTTP correctly, so the complete suite was rerun outside it.
