#pragma once
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <new>
#include <initializer_list>

// No CRT or C++ runtime: storage and text handling use Windows directly.
inline void* allocate(SIZE_T bytes)
{
    void* p = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, bytes ? bytes : 1);
    if (!p)
    {
        MessageBoxW(nullptr, L"Недостаточно памяти.", L"Stardisk's Player", MB_ICONERROR);
        ExitProcess(8);
    }
    return p;
}
inline void freeMemory(void* p)
{
    if (p)
        HeapFree(GetProcessHeap(), 0, p);
}
inline int clamp(int n, int lo, int hi)
{
    return n < lo ? lo : n > hi ? hi : n;
}

class Text
{
    wchar_t* data_ = nullptr;

  public:
    Text() = default;
    Text(const wchar_t* s)
    {
        assign(s);
    }
    Text(const Text& s)
    {
        assign(s.c_str());
    }
    Text(Text&& s) noexcept : data_(s.data_)
    {
        s.data_ = nullptr;
    }
    ~Text()
    {
        freeMemory(data_);
    }
    Text& operator=(const Text& s)
    {
        if (this != &s)
            assign(s.c_str());
        return *this;
    }
    Text& operator=(Text&& s) noexcept
    {
        if (this != &s)
        {
            freeMemory(data_);
            data_ = s.data_;
            s.data_ = nullptr;
        }
        return *this;
    }
    Text& operator=(const wchar_t* s)
    {
        assign(s);
        return *this;
    }
    const wchar_t* c_str() const
    {
        return data_ ? data_ : L"";
    }
    operator const wchar_t*() const
    {
        return c_str();
    }
    int size() const
    {
        return lstrlenW(c_str());
    }
    bool empty() const
    {
        return !*c_str();
    }
    void assign(const wchar_t* s, int n = -1)
    {
        if (!s)
            s = L"";
        if (n < 0)
            n = lstrlenW(s);
        auto p = static_cast<wchar_t*>(allocate((SIZE_T(n) + 1) * sizeof(wchar_t)));
        for (int i = 0; i < n; ++i)
            p[i] = s[i];
        freeMemory(data_);
        data_ = p;
    }
    Text& operator+=(const wchar_t* s)
    {
        int a = size(), b = lstrlenW(s);
        auto p = static_cast<wchar_t*>(allocate((SIZE_T(a) + b + 1) * sizeof(wchar_t)));
        for (int i = 0; i < a; ++i)
            p[i] = c_str()[i];
        for (int i = 0; i < b; ++i)
            p[a + i] = s[i];
        freeMemory(data_);
        data_ = p;
        return *this;
    }
    Text part(int first, int count = -1) const
    {
        first = clamp(first, 0, size());
        Text t;
        t.assign(c_str() + first, count < 0 ? size() - first : clamp(count, 0, size() - first));
        return t;
    }
};
inline Text operator+(const Text& a, const wchar_t* b)
{
    Text r(a);
    r += b;
    return r;
}
inline Text operator+(const wchar_t* a, const Text& b)
{
    Text r(a);
    r += b;
    return r;
}
inline Text operator+(const Text& a, const Text& b)
{
    Text r(a);
    r += b;
    return r;
}
inline bool equal(const wchar_t* a, const wchar_t* b)
{
    return lstrcmpW(a, b) == 0;
}
inline bool equalNoCase(const wchar_t* a, const wchar_t* b)
{
    return lstrcmpiW(a, b) == 0;
}
inline Text number(LONGLONG value)
{
    wchar_t buf[32];
    int i = 31;
    buf[i] = 0;
    bool minus = value < 0;
    ULONGLONG n = minus ? ULONGLONG(-(value + 1)) + 1 : ULONGLONG(value);
    do
    {
        buf[--i] = wchar_t(L'0' + n % 10);
        n /= 10;
    } while (n);
    if (minus)
        buf[--i] = L'-';
    return Text(buf + i);
}
inline bool integer(const wchar_t* s, int& n)
{
    while (*s == L' ')
        ++s;
    bool minus = *s == L'-';
    if (*s == L'-' || *s == L'+')
        ++s;
    if (*s < L'0' || *s > L'9')
        return false;
    unsigned value = 0;
    while (*s >= L'0' && *s <= L'9')
    {
        if (value > 214748364u)
            return false;
        value = value * 10 + (*s++ - L'0');
    }
    while (*s == L' ')
        ++s;
    if (*s || value > 2147483647u)
        return false;
    n = minus ? -int(value) : int(value);
    return true;
}
inline Text hexCode(DWORD n)
{
    wchar_t b[11] = L"0x00000000";
    for (int i = 9; i > 1; --i)
    {
        b[i] = L"0123456789ABCDEF"[n & 15];
        n >>= 4;
    }
    return b;
}
inline Text trim(const Text& t)
{
    int a = 0, b = t.size();
    while (a < b && t.c_str()[a] <= 32)
        ++a;
    while (b > a && t.c_str()[b - 1] <= 32)
        --b;
    return t.part(a, b - a);
}
inline Text lower(const Text& t)
{
    Text r(t);
    if (!r.empty())
        CharLowerBuffW(const_cast<wchar_t*>(r.c_str()), r.size());
    return r;
}
inline int lastSlash(const Text& t)
{
    for (int i = t.size() - 1; i >= 0; --i)
        if (t.c_str()[i] == L'\\')
            return i;
    return -1;
}
inline Text leaf(const Text& t)
{
    return t.part(lastSlash(t) + 1);
}
inline Text parent(const Text& t)
{
    int i = lastSlash(t);
    if (i == 2 && t.c_str()[1] == L':')
        ++i;
    return t.part(0, i < 0 ? 0 : i);
}
inline Text join(const Text& a, const Text& b)
{
    return a.empty() || a.c_str()[a.size() - 1] == L'\\' ? a + b : a + L"\\" + b;
}
inline bool exists(const wchar_t* p, bool folder = false)
{
    DWORD a = GetFileAttributesW(p);
    return a != INVALID_FILE_ATTRIBUTES && !!(a & FILE_ATTRIBUTE_DIRECTORY) == folder;
}
inline bool streamURL(const Text& p)
{
    Text s = lower(trim(p));
    return equal(s.part(0, 7), L"http://") || equal(s.part(0, 8), L"https://");
}
inline Text windowText(HWND h)
{
    int n = GetWindowTextLengthW(h);
    auto p = static_cast<wchar_t*>(allocate((SIZE_T(n) + 1) * 2));
    GetWindowTextW(h, p, n + 1);
    Text r(p);
    freeMemory(p);
    return r;
}
inline Text windowsError(DWORD code)
{
    wchar_t* p = nullptr;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr,
                   code, 0, (LPWSTR)&p, 0, nullptr);
    Text r = p ? trim(Text(p)) : Text(L"Ошибка Windows: ") + number(code);
    if (p)
        LocalFree(p);
    return r;
}
inline Text timeText(int seconds)
{
    if (seconds < 0)
        seconds = 0;
    return (seconds / 60 < 10 ? Text(L"0") : Text()) + number(seconds / 60) + L":" +
           (seconds % 60 < 10 ? Text(L"0") : Text()) + number(seconds % 60);
}
inline void pumpMessages()
{
    MSG m;
    while (PeekMessageW(&m, nullptr, 0, 0, PM_REMOVE))
    {
        if (m.message == WM_QUIT)
        {
            PostQuitMessage((int)m.wParam);
            break;
        }
        TranslateMessage(&m);
        DispatchMessageW(&m);
    }
}

class TextList
{
    Text** data_ = nullptr;
    int capacity_ = 0;

  public:
    int count = 0;
    ~TextList()
    {
        clear();
        freeMemory(data_);
    }
    const Text& operator[](int i) const
    {
        return *data_[i];
    }
    void clear()
    {
        for (int i = 0; i < count; ++i)
        {
            data_[i]->~Text();
            freeMemory(data_[i]);
        }
        count = 0;
    }
    void add(const Text& s)
    {
        if (count == capacity_)
        {
            int cap = capacity_ ? capacity_ * 2 : 16;
            auto p = static_cast<Text**>(allocate(SIZE_T(cap) * sizeof(Text*)));
            for (int i = 0; i < count; ++i)
                p[i] = data_[i];
            freeMemory(data_);
            data_ = p;
            capacity_ = cap;
        }
        data_[count++] = new (allocate(sizeof(Text))) Text(s);
    }
    bool contains(const Text& s) const
    {
        for (int i = 0; i < count; ++i)
            if (equal(*data_[i], s))
                return true;
        return false;
    }
};

inline bool ensureINI(const Text& path)
{
    HANDLE h = CreateFileW(path, GENERIC_WRITE, 0, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE)
        return false;
    LARGE_INTEGER size{};
    bool ok = !!GetFileSizeEx(h, &size);
    if (ok && !size.QuadPart)
    {
        wchar_t bom = 0xFEFF;
        DWORD n = 0;
        ok = WriteFile(h, &bom, 2, &n, nullptr) && n == 2;
    }
    CloseHandle(h);
    return ok;
}
inline Text readINI(const Text& path, const wchar_t* section, const wchar_t* key)
{
    auto p = static_cast<wchar_t*>(allocate(65536));
    GetPrivateProfileStringW(section, key, L"", p, 32768, path);
    Text r(p);
    freeMemory(p);
    return r;
}
inline bool writeINI(const Text& path, const wchar_t* section, const wchar_t* key, const wchar_t* value)
{
    return ensureINI(path) && WritePrivateProfileStringW(section, key, value, path);
}
Text encodeStream(const Text& value);
Text decodeStream(const Text& value);
