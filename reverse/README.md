# reverse.dll

**English** | [Русский](README.ru.md)

`reverse.dll` is an optional library that enables reverse playback in Stardisk's Player. The C++ and twinBASIC players load the same native x64 DLL from beside their EXE. No ActiveX/COM registration, installer or administrator rights are needed to load it.

Without the DLL, normal playback still works and the reverse menu item is hidden. Restart the player after adding the DLL: availability is checked when the player first requests it.

## Sources and build

- `reverse.cpp` — exported API, sessions, worker thread and playback control.
- `engine.h` — decoding, MP3 indexing, PCM reversal and waveOut output.
- `api.h` — shared C ABI declarations and the C++ dynamic loader.
- `build.ps1` — x64 DLL build using LLVM-MinGW.
- `tests.cpp`, `test.ps1` — API tests with a generated silent WAV fixture.

Only sources and build/test scripts are stored in Git. Compiled DLL/EXE files, test audio, logs and intermediate files are excluded.

On Windows 10/11 x64, with LLVM-MinGW installed, run from the repository root:

```powershell
.\reverse\build.ps1
```

The default toolchain is `C:\Program Files\LLVM-MinGW`. Another location can be selected with `-Toolchain 'C:\Tools\llvm-mingw'`. The result is `reverse/reverse.dll`; the build does not copy it into other folders.

The library uses system Media Foundation and waveOut. The C++ standard library is linked statically, so no third-party C++ runtime DLL needs to be distributed beside it.

## Enable reverse playback

Build the DLL and put it beside the player executable. For the repository's default build locations:

```powershell
# C++ player
Copy-Item .\reverse\reverse.dll .\cpp\reverse.dll

# twinBASIC player, after building its x64 EXE
Copy-Item .\reverse\reverse.dll .\twinbasic\reverse.dll
```

For a player built to another folder, copy the DLL to that folder instead. Restart the player, then use the reverse item in the Mode menu or right-click Play in the main or compact window.

The player's position and paused state are preserved when changing direction. New local tracks start from the end in reverse mode. Whole-track and A–B repeat follow the playback direction. Reverse playback runs at x1 with speed adjustment disabled; returning to a backend that supports speed restores the previous setting. HTTP/HTTPS streams play forward and cannot be reversed.

## API and limitations

The interface version is `0x00010000`. Exports use `WINAPI`, UTF-16 strings and an opaque session pointer (`LongPtr` in twinBASIC). The C++ player includes this folder's `api.h` directly. Callers must stop using a session before closing it.

| Exports, prefixed with `Reverse_` | Purpose |
| --- | --- |
| `Version` | Interface version |
| `OpenW`, `Close` | Open a local file and release the session |
| `Play`, `Pause`, `Stop` | Playback control |
| `Seek`, `Volume` | Original-file position in seconds; session volume from 0 to 100 |
| `Duration`, `Position`, `State`, `Format` | Duration, position, state, sample rate and channel count |
| `ErrorW` | Error text in a caller-provided buffer |
| `ExportW` | Explicit PCM WAV export for comparisons |

`OpenW` modes: 0 = forward with full decoding, 1 = reverse with full decoding, 2 = block reverse. The players use mode 2. States: 0 = stopped, 1 = playing, 2 = paused, 3 = ended, 4 = error.

Decoding and queued audio run in a worker thread. The source sample rate and channel count are preserved, with PCM16 samples. For MP3, the library keeps the compressed file in memory (up to 1 GiB), indexes frames and decodes sections with preceding context; block mode does not cache the entire decoded track. Other Media Foundation formats use timestamp seeking, whose accuracy depends on the system source. Mono and stereo have been checked; multichannel layouts have not been separately validated.

## Tests

```powershell
.\reverse\test.ps1
```

The script builds the DLL and tests, creates its own 30-second silent WAV under `reverse/qa/`, and checks the API, reverse position, pause, seeking, volume, stop, natural completion, released handles and startup without the DLL. It requires LLVM-MinGW, with the same `-Toolchain` option as the build script. Node.js is not needed. Test artifacts stay under ignored `build/` and `qa/` folders. Use a normal Windows session for system audio components.

To check integration with the C++ player, copy the built DLL beside `cpp/stardisk.exe` as above and run `cpp/test.ps1`. The test script also places that DLL beside its test executable. The earlier no-DLL checks remain relevant to normal standalone playback.

## Verification of this import

The DLL was built from these sources. All 25 API/missing-DLL checks passed, along with 119 native playback checks and 90 C++ application checks with the DLL installed. Integration covered direction switching, paused position, speed controls, track and A–B repeat, and streams. The twinBASIC project was not recompiled during this import.
