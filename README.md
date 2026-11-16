# Stardisk's Player

**English** | [Русский](README.ru.md)

_This project was originally created using Visual Basic 6.0 as a simple lightweight interface for Windows Media Player (wmp.dll) more than 10 years ago.
In 2022 I removed wmp.dll dependency, and replaced it with DirectShow.
Since then I've totally forgotten how to code in VB6 and this project has been abandoned._

_Now I have an AI (ChatGPT Codex)._
_First, it helped me port the project to twinBASIC, but later I decided to rewrite it in C++._
_ChatGPT rewrote the entire codebase into C++ and added a reverse function, which I have no idea how to implement._
_All three versions - VB6, twinBASIC, and C++ - are included in this repository)._

_So, I admit that this project contains a lot of AI code. You have been warned._

## What is it
A small audio player for files in the selected folder.
I don't need a music library or a pile of settings - I just want to play music from the current folder.
Also I like listening to music backwards (reverse) - it sounds so strange and exciting.

The **C++ version is the main version**, maintained on `main`, and the source for future player releases. The repository also preserves the earlier twinBASIC and VB6 versions.

## Downloads

[C++ v1.2.0 for Windows x64](https://github.com/Stardisk/Stardisk-s-player/releases/tag/v1.2.0):

- [stardisk.exe](https://github.com/Stardisk/Stardisk-s-player/releases/download/v1.2.0/stardisk.exe) — the player.
- [reverse.dll](https://github.com/Stardisk/Stardisk-s-player/releases/download/v1.2.0/reverse.dll) — optional reverse playback support. Put it beside the EXE and restart the player.

The [MIT license](https://github.com/Stardisk/Stardisk-s-player/releases/download/v1.2.0/LICENSE) is also attached to this release.

No installer is needed. Put the files in a writable folder; the player creates its INI settings there. The [historical VB6 release](https://github.com/Stardisk/Stardisk-s-player/releases/tag/StardiskPlayer) is still available.

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

## Optional reverse playback library

[reverse/](reverse/) contains the sources of **reverse.dll**, an optional x64 library that enables reverse playback in both the C++ and twinBASIC players. Build it with `.\reverse\build.ps1`, place the DLL beside the player EXE and restart the player. Use the Mode menu or right-click Play to switch direction. Without the DLL, normal playback still works and the reverse menu item stays hidden.

See the [library instructions](reverse/README.md) for build and activation steps, API and limitations. Compiled DLLs are excluded from Git; the repository contains their sources.

## Build from source

Install LLVM-MinGW. From PowerShell in the repository root:

```powershell
.\cpp\build.ps1
```

The default compiler location is `C:\Program Files\LLVM-MinGW`. Use `-Toolchain` to select another installation. The result is `cpp/stardisk.exe`.

See [C++ instructions](cpp/README.md) for requirements, build options and tests. Node.js and twinBASIC are not needed to build the C++ player.

## Releases

Player releases are built from `cpp/` on `main` at the chosen release commit or tag. Executables are distributed through [GitHub Releases](https://github.com/Stardisk/Stardisk-s-player/releases), not committed to the source tree. The existing VB6 release remains available as a historical version.

## Continuous integration

[Windows build](https://github.com/Stardisk/Stardisk-s-player/actions/workflows/windows-build.yml) checks the generated interface and compiles the C++ player and reverse DLL on pushes to `main`, pull requests and manual runs. It uses the pinned LLVM-MinGW distribution with SHA-256 verification. Successful runs provide an artifact containing the EXE, DLL, license, third-party notices and checksums. System audio tests are run locally; this workflow checks compilation and generation.

## Earlier versions

The [twinBASIC documentation](twinbasic/README.md) explains how to create the working `.twinproj` from its unpacked sources and export changes back to Git.

The two folders under `legacy/` preserve the original VB6 files, encodings and resources. Build output paths in their VBP files may need adjusting for another machine.

## License

[MIT License](LICENSE), copyright © 2026 Stardisk. Use it, modify it and share it, including commercially; keep the copyright and license notice with copies or substantial portions of the code.

Third-party copyright notices are preserved; see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
