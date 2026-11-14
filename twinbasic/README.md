# Stardisk's Player — twinBASIC

**English** | [Русский](README.ru.md)

This is the old VB6 player, converted to a twinBASIC project and updated into the current version. It started from the version without Windows Media Player in [legacy/vb6_nowmp](../legacy/vb6_nowmp/).

After the conversion, playback, file and stream handling, multimedia keys and the interface were updated. Unicode support was added, and the code was prepared for x64. The idea stayed the same: a small player that plays files from the selected folder without a music library or unnecessary settings.

Normal playback uses the system DirectShow and Media Foundation components. The code also supports an optional `reverse.dll`, but the library is not included in the repository yet. Without it, the reverse playback menu item is hidden; normal playback is designed to work without the DLL.

## What is stored in Git

- `source/` — the unpacked project: code, forms, embedded resources and settings.
- `tools/impexp.mjs` — a standalone twinBASIC project import/export tool, with its original MIT notice.
- `pack.ps1` — packs the sources into a working `.twinproj`.
- `export.ps1` — exports the saved project back into `source/`.

The working `.twinproj` and compiled EXE are excluded from Git. Code and form changes are compared using the files in `source/`.

### About the import/export tools

The scripts in this repository are a convenience workflow, not built-in Git integration. A standalone `impexp.mjs` tool is described in the [twinBASIC documentation](https://docs.twinbasic.com/Features/Packages/Import-Export-Tool). It reads and writes the project format itself, so compatibility can break after an IDE update. The officially supported way to read and write projects outside the IDE is to use `twinBASIC_win32.exe` or `twinBASIC_win64.exe` from the installation's `bin` folder.

`pack.ps1`, `export.ps1` and `tools/export-source.mjs` were added for this repository. They use the bundled standalone tool and require Node.js. They are not required by Git or GitHub.

You can also use the IDE's **Import from folder** command to open `source/`, and **File → Export Project** to export a saved project. IDE export empties its destination folder before writing: use a dedicated export folder, never the repository root or a folder containing the working `.twinproj`. IDE export may include compiler package sources; see the documentation before putting that export under version control.

## Opening and building with the bundled scripts

You need Windows, twinBASIC with Forms support, and Node.js 18 or newer. Node.js is only needed for packing and exporting with these scripts; it is not needed to run the player.

From the repository root:

```powershell
.\twinbasic\pack.ps1
```

Open `twinbasic/StardiskPlayer2NoWMP.twinproj` in the IDE. Select **win64** for the current version: the target architecture is stored in the IDE settings and does not travel with the project. References to twinBASIC packages and system libraries are listed in `source/Settings`; the installed compiler must have access to those dependencies.

Build the project in the IDE. The project settings place the EXE beside the `.twinproj`. Installing twinBASIC does not replace Windows codecs: supported formats depend on the system components available.

## Saving changes

1. Edit the project in the IDE and save it (`Ctrl+S`).
2. Run from the repository root:

   ```powershell
   .\twinbasic\export.ps1
   git diff -- twinbasic/source
   ```

3. Review the changes and commit them.

The export is first written to an empty temporary folder, then synchronized into `source/`, including the removal of files that are no longer in the project. Do not edit `source/` in parallel with the project in the IDE: exporting replaces its contents with the saved `.twinproj` contents.

After `git pull` or switching branches, close the project in the IDE and repack the updated sources:

```powershell
.\twinbasic\pack.ps1 -Force
```

`-Force` replaces the existing working `.twinproj`. Export any local changes you want to keep first. Without this option, the script refuses to overwrite the project.

## Verification of this import

Verification covers packing the sources and exporting them back, comparing the file list and contents. This checks how the project is stored in Git; compilation in twinBASIC and running without `reverse.dll` were not tested during this import.
