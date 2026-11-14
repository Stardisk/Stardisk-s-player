# Stardisk's Player

**English** | [Русский](README.ru.md)

A small audio player for files in the selected folder. I don't need a music library or a pile of settings — I just want to play music from the current folder.

The **C++ version is the main version**, maintained on `main`, and the source for future player releases. The repository also preserves the earlier twinBASIC and VB6 versions.

## Versions and development history

| Version | Sources | How it was made |
| --- | --- | --- |
| C++ — current | [cpp/](cpp/) | The twinBASIC version rewritten in modern C++20 with the help of AI. |
| twinBASIC | [twinbasic/](twinbasic/) | Converted from the legacy VB6 version without Windows Media Player, then updated; developed partly with vibe coding. |
| Legacy VB6 / WMP | [legacy/vb6/](legacy/vb6/) | Written without vibe coding. Preserved with the final local changes, project version 1.1.1. |
| Legacy VB6 / no WMP | [legacy/vb6_nowmp/](legacy/vb6_nowmp/) | Written without vibe coding. Uses DirectShow, project version 1.1.2. |

## Current C++ player

- Windows 10/11 x64, Russian interface and Unicode file paths.
- Plays files from the selected folder, with a compact window and multimedia keys.
- File rename, move and delete; HTTP/HTTPS streams.
- Shuffle, track repeat, A–B repeat, volume and playback speed controls.
- Uses system DirectShow and Media Foundation. Available formats depend on Windows codecs.
- Normal playback needs one EXE, with no Visual C++ Redistributable or .NET installation. Settings are saved beside it.

The optional reverse playback DLL will be handled separately. It is not included at this stage; normal playback works without it and the reverse menu item stays hidden.

## Build from source

Install LLVM-MinGW. From PowerShell in the repository root:

```powershell
.\cpp\build.ps1
```

The default compiler location is `C:\Program Files\LLVM-MinGW`. Use `-Toolchain` to select another installation. The result is `cpp/stardisk.exe`.

See [C++ instructions](cpp/README.md) for requirements, build options and tests. Node.js and twinBASIC are not needed to build the C++ player.

## Releases

Future player releases are built from `cpp/` on `main` at the chosen release commit or tag. Executables are distributed through [GitHub Releases](https://github.com/Stardisk/Stardisk-s-player/releases), not committed to the source tree. The existing VB6 release remains available as a historical version. Adding C++ sources does not publish a new release by itself.

## Earlier versions

The [twinBASIC documentation](twinbasic/README.md) explains how to create the working `.twinproj` from its unpacked sources and export changes back to Git.

The two folders under `legacy/` preserve the original VB6 files, encodings and resources. Build output paths in their VBP files may need adjusting for another machine.
