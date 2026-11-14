#pragma once
#include <windows.h>

// Native x64 C ABI. No COM registration, C++ objects, callbacks, or shared heaps.
// State: 0 stopped, 1 playing, 2 paused, 3 ended, 4 error.
// Mode: 0 forward/full decode, 1 reverse/full decode, 2 reverse/block decode.
#ifdef REVERSE_BUILD
#define REV extern "C" __declspec(dllexport)
REV DWORD WINAPI Reverse_Version();
REV HRESULT WINAPI Reverse_OpenW(const wchar_t*, int, double, void**);
REV void WINAPI Reverse_Close(void*);
REV HRESULT WINAPI Reverse_Play(void*);
REV HRESULT WINAPI Reverse_Pause(void*);
REV HRESULT WINAPI Reverse_Stop(void*);
REV HRESULT WINAPI Reverse_Seek(void*, double);
REV HRESULT WINAPI Reverse_Volume(void*, int);
REV double WINAPI Reverse_Duration(void*);
REV double WINAPI Reverse_Position(void*);
REV int WINAPI Reverse_State(void*);
REV HRESULT WINAPI Reverse_Format(void*, DWORD*, DWORD*);
REV void WINAPI Reverse_ErrorW(void*, wchar_t*, int);
REV HRESULT WINAPI Reverse_ExportW(void*, const wchar_t*, double);
#else
struct ReverseApi
{
    HMODULE module = nullptr;
    bool tried = false;
    DWORD (WINAPI* version)() = nullptr;
    HRESULT (WINAPI* open)(const wchar_t*, int, double, void**) = nullptr;
    void (WINAPI* close)(void*) = nullptr;
    HRESULT (WINAPI* play)(void*) = nullptr;
    HRESULT (WINAPI* pause)(void*) = nullptr;
    HRESULT (WINAPI* stop)(void*) = nullptr;
    HRESULT (WINAPI* seek)(void*, double) = nullptr;
    HRESULT (WINAPI* volume)(void*, int) = nullptr;
    double (WINAPI* duration)(void*) = nullptr;
    double (WINAPI* position)(void*) = nullptr;
    int (WINAPI* state)(void*) = nullptr;
    HRESULT (WINAPI* format)(void*, DWORD*, DWORD*) = nullptr;
    void (WINAPI* error)(void*, wchar_t*, int) = nullptr;
    HRESULT (WINAPI* exportW)(void*, const wchar_t*, double) = nullptr;
    bool load()
    {
        if (tried) return module != nullptr;
        tried = true;
        auto path = static_cast<wchar_t*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, 65536));
        if (!path) return false;
        DWORD n = GetModuleFileNameW(nullptr, path, 32768);
        if (n && n < 32755)
        {
            while (n && path[n-1] != L'\\') --n;
            lstrcpyW(path+n, L"reverse.dll");
            module = LoadLibraryExW(path, nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32);
        }
        HeapFree(GetProcessHeap(), 0, path);
        if (!module) return false;
#define GET(member, name) member = reinterpret_cast<decltype(member)>(GetProcAddress(module, name)); if (!member) { FreeLibrary(module); module = nullptr; return false; }
        GET(version, "Reverse_Version")
        GET(open, "Reverse_OpenW") GET(close, "Reverse_Close")
        GET(play, "Reverse_Play") GET(pause, "Reverse_Pause") GET(stop, "Reverse_Stop")
        GET(seek, "Reverse_Seek") GET(volume, "Reverse_Volume")
        GET(duration, "Reverse_Duration") GET(position, "Reverse_Position") GET(state, "Reverse_State")
        GET(format, "Reverse_Format") GET(error, "Reverse_ErrorW") GET(exportW, "Reverse_ExportW")
#undef GET
        if (version() != 0x10000) { FreeLibrary(module); module = nullptr; }
        return module != nullptr;
    }
    void unload()
    {
        if (module) FreeLibrary(module);
        module = nullptr; tried = false;
    }
};
#endif
