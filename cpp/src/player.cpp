#include "playback.h"
#include "ui.generated.h"
#include <shobjidl.h>
#include <uxtheme.h>

namespace
{
constexpr wchar_t title[] = L"Stardisk's Player";
constexpr UINT mediaMessage = WM_APP + 1, startupMessage = WM_APP + 2;
constexpr DWORD mainStyle = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
struct App
{
    HINSTANCE instance{};
    HWND main{}, compact{}, tooltip{};
    HWND controls[256]{};
    HFONT fonts[256]{};
    HMENU menu{}, streamsMenu{};
    ITaskbarList3* taskbar{};
    UINT taskbarButtonCreated{};
    HICON folderClosed{}, folderOpen{};
    HANDLE mutex{};
    Playback playback;
    Text home, presets, repeatlist, folder, current, info;
    TextList streams, directoryPaths;
    bool checked[256]{}, initializing = true, loading = false, fileBusy = false;
    bool isStream = false, repeating = false, dragging = false, ended = false;
    bool reverseEnabled = false;
    int savedForwardSpeed = 100;
    bool showCaption = false, closeCompact = false, minimizeCompact = false, showTaskbarProgress = false;
    int status = 0, volume = 100, speed = 100, start = 0, end = 0, compactStep = 5;
    double realPosition = 0;
    int scrollX = 0, scrollWidth = 0;
    HFONT compactTitleFont{};
    DWORD random = 1;
    bool registered[4]{};
    int lastCommand = 0, lastSource = 0;
    ULONGLONG lastTick = 0, moveStarted = 0;
}* app;

HWND control(int id)
{
    return app->controls[id];
}
const ControlDef* definition(int id)
{
    for (const auto& d : controlDefs)
        if (d.id == id)
            return &d;
    return nullptr;
}
void message(const Text& text, UINT flags = MB_ICONEXCLAMATION)
{
    MessageBoxW(IsWindowVisible(app->main) ? app->main : app->compact, text, title, flags);
}
void setText(int id, const Text& text)
{
    HWND h = control(id);
    const ControlDef* d = definition(id);
    if (!h || equal(windowText(h), text))
        return;
    if (d && d->autosize && d->kind == 0)
    {
        HDC dc = GetDC(h);
        HGDIOBJ old = SelectObject(dc, app->fonts[id]);
        SIZE size{};
        GetTextExtentPoint32W(dc, text, text.size(), &size);
        SelectObject(dc, old);
        ReleaseDC(h, dc);
        int x = d->x;
        if (d->align == 1)
            x += d->w - size.cx;
        else if (d->align == 2)
            x += (d->w - size.cx) / 2;
        SetWindowPos(h, nullptr, x, d->y, size.cx ? size.cx : 1, d->h, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    SetWindowTextW(h, text);
    InvalidateRect(h, nullptr, TRUE);
}
void persist(const wchar_t* section, const wchar_t* key, const Text& value)
{
    if (!writeINI(app->presets, section, key, value))
        message(Text(L"Не удалось сохранить настройки:\r\n") + windowsError(GetLastError()));
}
int setting(const wchar_t* section, const wchar_t* key, int fallback)
{
    int n;
    return integer(readINI(app->presets, section, key), n) ? n : fallback;
}
Text readStream(const wchar_t* key)
{
    return decodeStream(readINI(app->presets, L"Streams", key));
}
bool writeStream(const wchar_t* key, const Text& value)
{
    return writeINI(app->presets, L"Streams", key, encodeStream(value));
}
void invalidateButtons()
{
    const int ids[] = {main_cmdPlay,  main_cmdPause,   main_cmdStop,     main_chkRepeat, main_chkShuffle,
                       main_cmdApply, compact_cmdPlay, compact_cmdPause, compact_cmdStop};
    for (int id : ids)
        InvalidateRect(control(id), nullptr, TRUE);
}
double playbackFraction(double duration)
{
    if (!(duration > 0) || !(app->realPosition > 0))
        return 0;
    return app->realPosition >= duration ? 1 : app->realPosition / duration;
}
void updateTaskbarProgress()
{
    if (!app->taskbar || !app->showTaskbarProgress || !IsWindowVisible(app->main))
        return;
    if (!app->status || !app->playback.hasMedia())
    {
        app->taskbar->SetProgressState(app->main, TBPF_NOPROGRESS);
        return;
    }

    double duration = app->isStream ? 0 : app->playback.duration();
    TBPFLAG state = app->status == 2 ? TBPF_PAUSED : duration > 0 ? TBPF_NORMAL : TBPF_INDETERMINATE;
    app->taskbar->SetProgressState(app->main, state);
    if (duration > 0)
    {
        constexpr ULONGLONG scale = 10000;
        app->taskbar->SetProgressValue(app->main, ULONGLONG(playbackFraction(duration) * scale), scale);
    }
}
void changeStatus(int status)
{
    app->status = status;
    invalidateButtons();
    updateTaskbarProgress();
}
void setBrowserEnabled(bool enabled)
{
    for (int id : {main_Drive1, main_Dir1, main_File1})
        EnableWindow(control(id), enabled);
}
void streamControls()
{
    for (int id : {main_hsprogress, main_cmdApply, main_chkRepeat, main_txtStartTime, main_txtEndTime})
        EnableWindow(control(id), !app->isStream);
    for (int id : {menu_mnuDelete, menu_mnuRename, menu_mnuMove, menu_mnuOpenLocation})
        EnableMenuItem(app->menu, id, MF_BYCOMMAND | (app->isStream ? MF_GRAYED : MF_ENABLED));
    EnableMenuItem(app->menu, menu_mnuRepeat, MF_BYCOMMAND | (app->isStream ? MF_GRAYED : MF_ENABLED));
    EnableMenuItem(app->menu, menu_mnuReverse, MF_BYCOMMAND | (app->isStream ? MF_GRAYED : MF_ENABLED));
    setText(main_cmdPlay, app->reverseEnabled && !app->isStream ? L"3" : L"4");
    setText(compact_cmdPlay, app->reverseEnabled && !app->isStream ? L"3" : L"4");
}
bool blocked()
{
    return app->initializing || app->loading || app->fileBusy || app->playback.busy;
}
int fileCount()
{
    return (int)SendMessageW(control(main_File1), LB_GETCOUNT, 0, 0);
}
int selectedIndex()
{
    return (int)SendMessageW(control(main_File1), LB_GETCURSEL, 0, 0);
}
Text fileName(int index)
{
    HWND h = control(main_File1);
    LRESULT n = SendMessageW(h, LB_GETTEXTLEN, index, 0);
    if (n == LB_ERR)
        return Text();
    auto p = static_cast<wchar_t*>(allocate((SIZE_T(n) + 1) * 2));
    SendMessageW(h, LB_GETTEXT, index, (LPARAM)p);
    Text result(p);
    freeMemory(p);
    return result;
}
Text selectedFile()
{
    Text s = fileName(selectedIndex());
    return s.empty() ? Text() : join(app->folder, s);
}
Text playingCaption()
{
    return app->isStream ? app->current : leaf(app->current);
}
void updateCaption()
{
    SetWindowTextW(app->main, app->showCaption && app->status ? playingCaption().c_str() : title);
}
void timing()
{
    setText(main_lblTiming,
            timeText(int(app->realPosition)) +
                (app->isStream ? Text(L" / поток") : Text(L" / ") + timeText(int(app->playback.duration()))));
}
void seek(double seconds)
{
    if (app->isStream || !app->playback.hasMedia() || app->playback.busy)
        return;
    double length = app->playback.duration();
    if (seconds < 0)
        seconds = 0;
    if (seconds > length)
        seconds = length;
    HRESULT hr = app->playback.seek(seconds);
    if (FAILED(hr))
    {
        message(Text(L"Не удалось изменить позицию воспроизведения:\r\n") + hexCode(hr) + L"\r\n" + windowsError(hr));
        return;
    }
    app->realPosition = seconds;
    SendMessageW(control(main_hsprogress), TBM_SETPOS, TRUE, (LPARAM)seconds);
    timing();
    updateTaskbarProgress();
}
bool handlePositionKey(const MSG& msg)
{
    if ((msg.message != WM_KEYDOWN && msg.message != WM_KEYUP) || (msg.wParam != VK_LEFT && msg.wParam != VK_RIGHT))
        return false;
    if (!IsWindowEnabled(app->main) || !(msg.hwnd == app->main || IsChild(app->main, msg.hwnd) ||
                                         msg.hwnd == app->compact || IsChild(app->compact, msg.hwnd)))
        return false;
    // Time fields keep normal caret movement. Other controls must not use
    // these keys for dialog focus navigation or drive/file selection.
    if (msg.hwnd == control(main_txtStartTime) || msg.hwnd == control(main_txtEndTime))
        return false;
    if (msg.message == WM_KEYDOWN && !blocked() && !app->isStream && app->playback.hasMedia())
    {
        app->dragging = false;
        // Work in absolute media time, regardless of playback direction.
        seek(app->playback.position() + (msg.wParam == VK_LEFT ? -10.0 : 10.0));
    }
    return true;
}
void volume(int value)
{
    app->volume = clamp(value, 0, 100);
    SetScrollPos(control(main_hsvolume), SB_CTL, app->volume, TRUE);
    if (app->playback.hasMedia())
        app->playback.volume(app->volume);
    setText(main_lblVolume, number(app->volume));
    setText(compact_lblVolume, number(app->volume));
    if (!app->initializing)
        persist(L"General", L"Volume", number(app->volume));
}
Text rateText(int n)
{
    Text s = Text(L"x") + number(n / 100);
    if (n % 100)
    {
        s += L".";
        s += number((n % 100) / 10);
        if (n % 10)
            s += number(n % 10);
    }
    return s;
}
void speed(int value)
{
    if (!IsWindowEnabled(control(main_hsRate)) || (app->playback.hasMedia() && !app->playback.supportsSpeed))
        return;
    app->speed = clamp(value, 50, 200);
    SetScrollPos(control(main_hsRate), SB_CTL, app->speed, TRUE);
    setText(main_lblRate, rateText(app->speed));
    if (app->playback.hasMedia() && app->playback.supportsSpeed)
    {
        app->playback.rate(app->speed);
        app->realPosition = app->playback.position();
        timing();
    }
}
void enableSpeedControl(bool enabled)
{
    HWND scrollbar = control(main_hsRate);
    EnableWindow(scrollbar, enabled);
    EnableScrollBar(scrollbar, SB_CTL, enabled ? ESB_ENABLE_BOTH : ESB_DISABLE_BOTH);
    InvalidateRect(scrollbar, nullptr, TRUE);
}
void clearMedia(bool clearCurrent)
{
    KillTimer(app->main, 1);
    app->playback.close();
    updateTaskbarProgress();
    app->dragging = false;
    app->repeating = false;
    app->checked[main_cmdApply] = false;
    app->ended = false;
    app->realPosition = 0;
    setText(main_lblBitrate, L"");
    setText(main_lblSysInfo, L"");
    app->info = L"";
    if (clearCurrent)
    {
        app->current = L"";
        app->isStream = false;
        changeStatus(0);
        setText(main_lblTiming, L"--:-- / --:--");
        setText(main_lblSize, L"");
        setText(main_lblItems, L"");
        updateCaption();
    }
    streamControls();
    enableSpeedControl(!app->isStream && !app->reverseEnabled);
    invalidateButtons();
}
void playbackFailure(HRESULT hr, const Text& attempts)
{
    Text file(app->current);
    clearMedia(false);
    changeStatus(0);
    Text explanation;
    switch (DWORD(hr))
    {
    case 0x80040266:
        explanation = L"DirectShow не смог соединить компоненты воспроизведения: они используют несовместимые способы "
                      L"передачи данных.";
        break;
    case 0x80040218:
        explanation = L"DirectShow не смог подобрать цепочку компонентов для воспроизведения этого файла.";
        break;
    case 0x80040265:
    case 0x8004025C:
        explanation = L"Выбранная цепочка воспроизведения не поддерживает формат звука в этом файле.";
        break;
    case 0x80070002:
    case 0x80070003:
        explanation = L"Файл или папка не найдены. Возможно, файл был перемещён или удалён.";
        break;
    case 0x80070005:
        explanation = L"Нет доступа к файлу или одному из компонентов воспроизведения.";
        break;
    case 0x8007007E:
        explanation = L"Не удалось загрузить необходимый системный компонент или одну из его зависимостей.";
        break;
    default:
        explanation = L"При подготовке воспроизведения произошла ошибка. Подробности приведены ниже.";
    }
    Text result = Text(L"Не удалось воспроизвести файл:\r\n") + file + L"\r\n\r\n" + explanation +
                  L"\r\n\r\nТехнические сведения:\r\nКод: " + hexCode(hr) + L" (" + number(LONG(hr)) + L")\r\n" +
                  windowsError(hr);
    if (!attempts.empty())
        result += Text(L"\r\n\r\nПопытки открытия:\r\n") + attempts;
    message(result);
}
void loadFile(const Text& path)
{
    if (app->loading || app->fileBusy || app->playback.busy || path.empty())
        return;
    Text file(path);
    app->loading = true;
    setBrowserEnabled(false);
    clearMedia(false);
    app->current = file;
    app->isStream = streamURL(file);
    streamControls();
    Text attempts;
    HRESULT hr = app->playback.open(file, attempts, app->reverseEnabled && !app->isStream);
    if (SUCCEEDED(hr))
    {
        if (!app->playback.supportsSpeed)
            app->speed = 100;
        setText(main_lblRate, rateText(app->speed));
        SetScrollPos(control(main_hsRate), SB_CTL, app->speed, TRUE);
        enableSpeedControl(app->playback.supportsSpeed);
        app->playback.volume(app->volume);
        app->playback.rate(app->speed);
        if (!app->isStream)
            hr = app->playback.seek(app->playback.isReverse() ? app->playback.duration() : 0);
        if (SUCCEEDED(hr))
            hr = app->playback.run();
    }
    if (FAILED(hr))
    {
        app->loading = false;
        setBrowserEnabled(true);
        playbackFailure(hr, attempts);
        return;
    }
    app->info =
        app->isStream ? Text(app->playback.component) : trackDisplay(file) + L"\r\n[" + app->playback.component + L"]";
    setText(main_lblSysInfo, app->info);
    app->realPosition = app->playback.isReverse() ? app->playback.duration() : 0;
    changeStatus(1);
    updateCaption();
    if (app->isStream)
    {
        setText(main_lblItems, L"Поток");
        setText(main_lblSize, L"");
        setText(main_lblBitrate, L"");
        SendMessageW(control(main_hsprogress), TBM_SETRANGEMAX, TRUE, 1);
        SendMessageW(control(main_hsprogress), TBM_SETPOS, TRUE, (LPARAM)app->realPosition);
    }
    else
    {
        setText(main_lblItems, number(selectedIndex() + 1) + L" / " + number(fileCount()));
        WIN32_FILE_ATTRIBUTE_DATA data{};
        GetFileAttributesExW(file, GetFileExInfoStandard, &data);
        ULONGLONG bytes = (ULONGLONG(data.nFileSizeHigh) << 32) | data.nFileSizeLow;
        ULONGLONG hundredths = (bytes * 100 + 524288) / 1048576;
        setText(main_lblSize, number(hundredths / 100) + L"." + (hundredths % 100 < 10 ? Text(L"0") : Text()) +
                                  number(hundredths % 100) + L" МБ");
        double length = app->playback.duration();
        setText(main_lblBitrate, length > 0
                                     ? Text(L"~") + number(LONGLONG(double(bytes) * 8.0 / length / 1000.0)) + L" кбит/с"
                                     : Text());
        SendMessageW(control(main_hsprogress), TBM_SETRANGEMAX, TRUE, (LPARAM)length);
        SendMessageW(control(main_hsprogress), TBM_SETPOS, TRUE, (LPARAM)app->realPosition);
        Text key = lower(file);
        int n;
        app->start = integer(readINI(app->repeatlist, key, L"Start"), n) ? n : 0;
        app->end = integer(readINI(app->repeatlist, key, L"End"), n) ? n : 0;
        setText(main_txtStartTime, timeText(app->start));
        setText(main_txtEndTime, timeText(app->end));
        persist(L"General", L"LastIndex", number(selectedIndex()));
    }
    timing();
    SetTimer(app->main, 1, app->playback.isReverse() ? 100 : 1000, nullptr);
    app->loading = false;
    setBrowserEnabled(true);
}
void selectTrack(int index)
{
    int count = fileCount();
    if (!count)
    {
        if (!app->initializing)
            clearMedia(true);
        return;
    }
    index = clamp(index, 0, count - 1);
    SendMessageW(control(main_File1), LB_SETCURSEL, index, 0);
    if (!app->initializing)
        loadFile(selectedFile());
}
void nextTrack(bool previous = false)
{
    if (blocked() || !fileCount())
        return;
    int index = selectedIndex();
    if (previous)
        index = index <= 0 ? fileCount() - 1 : index - 1;
    else if (app->checked[main_chkShuffle])
    {
        app->random ^= app->random << 13;
        app->random ^= app->random >> 17;
        app->random ^= app->random << 5;
        index = app->random % fileCount();
    }
    else
    {
        ++index;
        if (index >= fileCount())
            index = 0;
    }
    selectTrack(index);
}
void play()
{
    if (blocked())
        return;
    if (!app->playback.hasMedia())
    {
        if (app->isStream)
            loadFile(app->current);
        else
            selectTrack(selectedIndex() < 0 ? 0 : selectedIndex());
        return;
    }
    if (app->repeating)
        seek(app->playback.isReverse() ? (app->end ? app->end : app->playback.duration()) : app->start);
    HRESULT hr = app->playback.run();
    if (FAILED(hr))
    {
        playbackFailure(hr, Text());
        return;
    }
    changeStatus(1);
    updateCaption();
}
void pause()
{
    if (blocked() || !app->playback.hasMedia())
        return;
    HRESULT hr = app->playback.pause();
    if (FAILED(hr))
    {
        message(windowsError(hr));
        return;
    }
    if (!app->isStream)
    {
        app->realPosition = app->playback.position();
        SendMessageW(control(main_hsprogress), TBM_SETPOS, TRUE, (LPARAM)app->realPosition);
        timing();
    }
    changeStatus(2);
}
void stop()
{
    if (blocked() || !app->playback.hasMedia())
        return;
    if (app->isStream)
    {
        clearMedia(false);
        setText(main_lblTiming, L"00:00 / поток");
    }
    else
    {
        app->playback.stop();
        seek(app->playback.isReverse() ? app->playback.duration() : 0);
    }
    changeStatus(0);
    updateCaption();
}
bool audioFile(const Text& file)
{
    Text s = lower(file);
    for (const wchar_t* ext : {L".mp3", L".wma", L".ogg", L".wav"})
        if (equal(s.part(s.size() - 4), ext))
            return true;
    return false;
}
void refreshFiles()
{
    HWND list = control(main_File1);
    SendMessageW(list, WM_SETREDRAW, FALSE, 0);
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    auto data = static_cast<WIN32_FIND_DATAW*>(allocate(sizeof(WIN32_FIND_DATAW)));
    HANDLE find = FindFirstFileW(join(app->folder, L"*"), data);
    if (find != INVALID_HANDLE_VALUE)
    {
        do
        {
            if (!(data->dwFileAttributes &
                  (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) &&
                audioFile(Text(data->cFileName)))
                SendMessageW(list, LB_ADDSTRING, 0, (LPARAM)data->cFileName);
        } while (FindNextFileW(find, data));
        FindClose(find);
    }
    freeMemory(data);
    SendMessageW(list, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(list, nullptr, TRUE);
}
void refreshDirectories()
{
    HWND list = control(main_Dir1);
    SendMessageW(list, WM_SETREDRAW, FALSE, 0);
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    app->directoryPaths.clear();
    Text current = app->folder;
    TextList ancestry;
    for (;;)
    {
        ancestry.add(current);
        Text p = parent(current);
        if (p.empty() || equalNoCase(p, current))
            break;
        current = p;
    }
    for (int i = ancestry.count - 1; i >= 0; --i)
    {
        app->directoryPaths.add(ancestry[i]);
        Text display = leaf(ancestry[i]);
        if (display.empty())
            display = ancestry[i];
        SendMessageW(list, LB_ADDSTRING, 0, (LPARAM)display.c_str());
        SendMessageW(list, LB_SETITEMDATA, app->directoryPaths.count - 1, (LPARAM)((ancestry.count - i - 1) << 1 | 1));
    }
    int currentRow = app->directoryPaths.count - 1;
    // Sort child directory names with the same Windows text comparison as the file list.
    HWND temp = CreateWindowW(L"LISTBOX", L"", LBS_SORT, 0, 0, 0, 0, app->main, nullptr, app->instance, nullptr);
    auto data = static_cast<WIN32_FIND_DATAW*>(allocate(sizeof(WIN32_FIND_DATAW)));
    HANDLE find = FindFirstFileW(join(app->folder, L"*"), data);
    if (find != INVALID_HANDLE_VALUE)
    {
        do
        {
            if ((data->dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) &&
                !(data->dwFileAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) &&
                !equal(data->cFileName, L".") && !equal(data->cFileName, L".."))
                SendMessageW(temp, LB_ADDSTRING, 0, (LPARAM)data->cFileName);
        } while (FindNextFileW(find, data));
        FindClose(find);
    }
    freeMemory(data);
    int count = (int)SendMessageW(temp, LB_GETCOUNT, 0, 0);
    auto buf = static_cast<wchar_t*>(allocate(65536));
    for (int i = 0; i < count; ++i)
    {
        SendMessageW(temp, LB_GETTEXT, i, (LPARAM)buf);
        app->directoryPaths.add(join(app->folder, Text(buf)));
        int row = (int)SendMessageW(list, LB_ADDSTRING, 0, (LPARAM)buf);
        SendMessageW(list, LB_SETITEMDATA, row, (LPARAM)(ancestry.count << 1));
    }
    freeMemory(buf);
    DestroyWindow(temp);
    SendMessageW(list, LB_SETCURSEL, currentRow, 0);
    SendMessageW(list, LB_SETTOPINDEX, currentRow, 0);
    SendMessageW(list, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(list, nullptr, TRUE);
}
bool setFolder(const Text& folder, bool start = true)
{
    if (!exists(folder, true))
    {
        message(Text(L"Папка не найдена или недоступна: ") + folder);
        return false;
    }
    auto path = static_cast<wchar_t*>(allocate(65536));
    DWORD n = GetFullPathNameW(folder, 32768, path, nullptr);
    if (!n || n >= 32768)
    {
        freeMemory(path);
        message(windowsError(GetLastError()));
        return false;
    }
    app->folder = path;
    freeMemory(path);
    while (app->folder.size() > 3 && app->folder.c_str()[app->folder.size() - 1] == L'\\')
        app->folder = app->folder.part(0, app->folder.size() - 1);
    refreshDirectories();
    refreshFiles();
    if (app->folder.size() >= 2 && app->folder.c_str()[1] == L':')
    {
        wchar_t c = app->folder.c_str()[0];
        if (c >= L'a' && c <= L'z')
            c -= 32;
        int count = (int)SendMessageW(control(main_Drive1), CB_GETCOUNT, 0, 0);
        for (int i = 0; i < count; ++i)
            if (SendMessageW(control(main_Drive1), CB_GETITEMDATA, i, 0) == c)
                SendMessageW(control(main_Drive1), CB_SETCURSEL, i, 0);
    }
    persist(L"General", L"LastPath", app->folder);
    if (start)
        selectTrack(0);
    return true;
}
void refreshStreams()
{
    while (GetMenuItemCount(app->streamsMenu) > 0)
        DeleteMenu(app->streamsMenu, 0, MF_BYPOSITION);
    if (!app->streams.count)
        AppendMenuW(app->streamsMenu, MF_STRING | MF_GRAYED, 0, L"(Нет сохранённых потоков)");
    else
    {
        for (int i = 0; i < app->streams.count; ++i)
        {
            Text cap;
            for (int k = 0; k < app->streams[i].size(); ++k)
            {
                wchar_t b[2] = {app->streams[i].c_str()[k], 0};
                cap += b;
                if (b[0] == L'&')
                    cap += L"&";
            }
            AppendMenuW(app->streamsMenu, MF_STRING, menu_streamBase + i, cap);
        }
        AppendMenuW(app->streamsMenu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(app->streamsMenu, MF_STRING, menu_AddStream, L"Добавить поток...\tCtrl+L");
}
bool saveStreams()
{
    for (int i = 0; i < app->streams.count; ++i)
        if (!writeStream(Text(L"URL") + number(i + 1), app->streams[i]))
            return false;
    return writeStream(L"Count", number(app->streams.count));
}
void loadStreams()
{
    Text countText = readStream(L"Count");
    int count = 0;
    integer(countText, count);
    count = clamp(count, 0, 32767);
    for (int i = 0; i < count; ++i)
    {
        Text address = trim(readStream(Text(L"URL") + number(i + 1)));
        if (streamURL(address) && !app->streams.contains(address))
            app->streams.add(address);
    }
    if (countText.empty())
    {
        Text address = readINI(app->presets, L"General", L"LastStreamURL");
        if (streamURL(address))
        {
            app->streams.add(address);
            saveStreams();
        }
    }
    refreshStreams();
}

struct Input
{
    const wchar_t* prompt;
    const wchar_t* caption;
    const wchar_t* initial;
    Text result;
    bool accepted = false;
};
INT_PTR CALLBACK inputProc(HWND h, UINT msg, WPARAM w, LPARAM l)
{
    auto input = reinterpret_cast<Input*>(GetWindowLongPtrW(h, DWLP_USER));
    if (msg == WM_INITDIALOG)
    {
        input = reinterpret_cast<Input*>(l);
        SetWindowLongPtrW(h, DWLP_USER, l);
        SetWindowTextW(h, input->caption);
        SetDlgItemTextW(h, 100, input->prompt);
        SetDlgItemTextW(h, 101, input->initial);
        SendDlgItemMessageW(h, 101, EM_SETLIMITTEXT, 32760, 0);
        SendDlgItemMessageW(h, 101, EM_SETSEL, 0, -1);
        SetFocus(GetDlgItem(h, 101));
        return FALSE;
    }
    if (msg == WM_COMMAND && (LOWORD(w) == IDOK || LOWORD(w) == IDCANCEL))
    {
        if (LOWORD(w) == IDOK)
        {
            input->result = windowText(GetDlgItem(h, 101));
            input->accepted = true;
        }
        EndDialog(h, LOWORD(w));
        return TRUE;
    }
    if (msg == WM_CLOSE)
    {
        EndDialog(h, IDCANCEL);
        return TRUE;
    }
    return FALSE;
}
Text inputBox(const Text& prompt, const Text& initial = Text(), const wchar_t* caption = L"")
{
    Input input{prompt, caption, initial, Text(), false};
    DialogBoxParamW(app->instance, MAKEINTRESOURCEW(2), app->main, inputProc, (LPARAM)&input);
    return input.accepted ? input.result : Text();
}
void addStream()
{
    Text address = readStream(L"LastURL");
    if (address.empty())
        address = readINI(app->presets, L"General", L"LastStreamURL");
    address = trim(inputBox(L"Введите HTTP или HTTPS URL аудиопотока:", address, L"Добавить поток"));
    if (address.empty())
        return;
    bool valid = streamURL(address);
    for (int i = 0; i < address.size(); ++i)
        if (address.c_str()[i] == L'\r' || address.c_str()[i] == L'\n' || address.c_str()[i] == L'\t')
            valid = false;
    if (!valid)
    {
        message(L"Укажите HTTP или HTTPS адрес потока одной строкой.");
        return;
    }
    if (!app->streams.contains(address))
    {
        if (app->streams.count >= 32767)
        {
            message(L"Список потоков заполнен.");
            return;
        }
        app->streams.add(address);
        if (!saveStreams())
        {
            message(L"Не удалось сохранить список потоков в presets.ini.");
            return;
        }
        refreshStreams();
    }
    writeStream(L"LastURL", address);
    loadFile(address);
    if (app->playback.hasMedia())
        persist(L"General", L"LastStreamURL", address);
}
void showCompact()
{
    app->compactStep = clamp(setting(L"General", L"CompactStepVolume", 5), 1, 100);
    int x = setting(L"CompactPosition", L"X", 0) / 15, y = setting(L"CompactPosition", L"Y", 0) / 15;
    SetWindowPos(app->compact, HWND_TOPMOST, x, y, 301, 17, SWP_SHOWWINDOW);
    ShowWindow(app->main, SW_HIDE);
    SetTimer(app->compact, 2, 500, nullptr);
}
void returnMain()
{
    RECT r{};
    GetWindowRect(app->compact, &r);
    persist(L"CompactPosition", L"X", number(r.left * 15));
    persist(L"CompactPosition", L"Y", number(r.top * 15));
    KillTimer(app->compact, 2);
    ShowWindow(app->compact, SW_HIDE);
    ShowWindow(app->main, SW_RESTORE);
    updateTaskbarProgress();
    SetForegroundWindow(app->main);
}
bool validFileName(const Text& s)
{
    if (s.empty() || equal(s, L".") || equal(s, L"..") || s.c_str()[s.size() - 1] == L'.' ||
        s.c_str()[s.size() - 1] == L' ')
        return false;
    for (int i = 0; i < s.size(); ++i)
    {
        wchar_t c = s.c_str()[i];
        if (c < 32)
            return false;
        for (const wchar_t* p = L"\\/:*?\"<>|"; *p; ++p)
            if (c == *p)
                return false;
    }
    return true;
}
DWORD CALLBACK moveProgress(LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, DWORD, DWORD, HANDLE, HANDLE,
                            LPVOID)
{
    pumpMessages();
    return GetTickCount64() - app->moveStarted > 30000 ? PROGRESS_CANCEL : PROGRESS_CONTINUE;
}
void fileOperation(int id)
{
    if (app->isStream)
        return;
    Text source = selectedFile();
    if (source.empty())
        return;
    int oldIndex = selectedIndex();
    Text target, op;
    if (id == menu_mnuDelete)
    {
        op = L"Удаление";
        if (MessageBoxW(app->main, Text(L"Удалить файл?\r\n") + source, title, MB_YESNO | MB_ICONQUESTION) != IDYES)
            return;
    }
    if (id == menu_mnuRename)
    {
        op = L"Переименование";
        Text name = inputBox(Text(L"Новое имя для файла ") + leaf(source), leaf(source));
        if (name.empty())
            return;
        if (!validFileName(name))
        {
            message(L"Введите имя файла без запрещённых символов и без точки или пробела в конце.");
            return;
        }
        target = join(parent(source), name);
        if (equal(source, target))
            return;
    }
    Text moveFolder;
    if (id == menu_mnuMove)
    {
        op = L"Перемещение";
        moveFolder = trim(inputBox(L"Куда перемещаем?", readINI(app->presets, L"General", L"LastMovePath")));
        if (moveFolder.empty())
            return;
        if (!exists(moveFolder, true))
        {
            message(Text(L"Папка не найдена или недоступна: ") + moveFolder);
            return;
        }
        target = join(moveFolder, leaf(source));
        if (equalNoCase(source, target))
            return;
    }
    if (!exists(source))
    {
        message(Text(L"Файл не найден: ") + source);
        return;
    }
    app->fileBusy = true;
    EnableWindow(app->main, FALSE);
    clearMedia(true);
    BOOL ok = FALSE;
    if (id == menu_mnuDelete)
        ok = DeleteFileW(source);
    else if (id == menu_mnuRename)
        ok = MoveFileExW(source, target, 0);
    else
    {
        app->moveStarted = GetTickCount64();
        ok = MoveFileWithProgressW(source, target, moveProgress, nullptr, MOVEFILE_COPY_ALLOWED);
    }
    DWORD error = ok ? 0 : GetLastError();
    if (ok && id == menu_mnuMove && (exists(source) || !exists(target)))
    {
        ok = FALSE;
        error = ERROR_WRITE_FAULT;
    }
    refreshFiles();
    EnableWindow(app->main, TRUE);
    app->fileBusy = false;
    if (!ok)
    {
        message(op + L" не выполнено:\r\n" + windowsError(error) + L"\r\n" + source);
        return;
    }
    if (id == menu_mnuRename)
    {
        for (int i = 0; i < fileCount(); ++i)
            if (equalNoCase(fileName(i), leaf(target)))
            {
                oldIndex = i;
                break;
            }
    }
    if (id == menu_mnuMove)
        persist(L"General", L"LastMovePath", moveFolder);
    selectTrack(oldIndex);
}
bool parseTime(const Text& value, int& seconds)
{
    if (value.size() != 5 || value.c_str()[2] != L':')
        return false;
    for (int i : {0, 1, 3, 4})
        if (value.c_str()[i] < L'0' || value.c_str()[i] > L'9')
            return false;
    seconds = (value.c_str()[0] - L'0') * 600 + (value.c_str()[1] - L'0') * 60 + (value.c_str()[3] - L'0') * 10 +
              (value.c_str()[4] - L'0');
    return true;
}
void applyRepeat()
{
    if (app->isStream || selectedIndex() < 0)
    {
        app->checked[main_cmdApply] = false;
        return;
    }
    if (app->checked[main_cmdApply])
    {
        int start, end;
        if (!parseTime(windowText(control(main_txtStartTime)), start) ||
            !parseTime(windowText(control(main_txtEndTime)), end))
        {
            message(L"Неверные данные.", MB_OK);
            app->checked[main_cmdApply] = false;
            return;
        }
        if (start >= end)
        {
            message(L"Время начала должно быть меньше времени окончания", MB_OK);
            app->checked[main_cmdApply] = false;
            return;
        }
        app->start = start;
        app->end = end;
        app->repeating = true;
        if (end)
            seek(app->playback.isReverse() ? (end ? end : app->playback.duration()) : start);
        EnableWindow(control(main_txtStartTime), FALSE);
        EnableWindow(control(main_txtEndTime), FALSE);
        if (start || end)
        {
            Text key = lower(app->current);
            if (!writeINI(app->repeatlist, key, L"Start", number(start)) ||
                !writeINI(app->repeatlist, key, L"End", number(end)))
                message(Text(L"Не удалось сохранить повтор:\r\n") + windowsError(GetLastError()));
        }
    }
    else
    {
        app->repeating = false;
        EnableWindow(control(main_txtStartTime), TRUE);
        EnableWindow(control(main_txtEndTime), TRUE);
    }
}
void toggleView(int id, bool& flag, const wchar_t* key)
{
    flag = !flag;
    CheckMenuItem(app->menu, id, MF_BYCOMMAND | (flag ? MF_CHECKED : MF_UNCHECKED));
    persist(L"View", key, flag ? L"1" : L"0");
}
void shutdown()
{
    if (blocked())
        return;
    for (int i = 0; i < 4; ++i)
        if (app->registered[i])
            UnregisterHotKey(app->main, 0x5110 + i);
    KillTimer(app->main, 1);
    KillTimer(app->compact, 2);
    app->playback.close();
    DestroyWindow(app->compact);
    DestroyWindow(app->main);
}
void toggleReverse()
{
    if (blocked() || app->isStream || !app->playback.reverseAvailable())
        return;
    bool enable = !app->reverseEnabled;
    int state = app->status;
    int targetSpeed = enable ? 100 : app->savedForwardSpeed;
    if (app->playback.hasMedia())
    {
        double position = app->playback.position();
        if (state == 1)
            app->playback.pause();
        app->loading = true;
        Playback candidate;
        Text attempts;
        HRESULT hr = candidate.open(app->current, attempts, enable);
        if (SUCCEEDED(hr))
            hr = candidate.volume(app->volume);
        if (SUCCEEDED(hr) && candidate.supportsSpeed)
            candidate.rate(targetSpeed);
        if (SUCCEEDED(hr))
            hr = candidate.seek(enable && state == 0 && position == 0 ? candidate.duration() : position);
        if (SUCCEEDED(hr) && state == 1)
            hr = candidate.run();
        if (FAILED(hr))
        {
            if (state == 1)
                app->playback.run();
            app->loading = false;
            message(Text(L"Не удалось изменить направление. Прежнее воспроизведение сохранено.\r\n") + attempts +
                    L"\r\n" + hexCode(hr));
            return;
        }
        app->playback.swap(candidate);
        app->realPosition = app->playback.position();
        if (!app->playback.supportsSpeed)
            targetSpeed = 100;
        app->info = trackDisplay(app->current) + L"\r\n[" + app->playback.component + L"]";
        setText(main_lblSysInfo, app->info);
        app->ended = false;
        SetTimer(app->main, 1, enable ? 100 : 1000, nullptr);
        app->loading = false;
    }
    if (enable)
        app->savedForwardSpeed = app->speed;
    app->speed = targetSpeed;
    app->reverseEnabled = enable;
    setText(main_lblRate, rateText(app->speed));
    SetScrollPos(control(main_hsRate), SB_CTL, app->speed, TRUE);
    enableSpeedControl(!enable && (!app->playback.hasMedia() || app->playback.supportsSpeed));
    CheckMenuItem(app->menu, menu_mnuReverse, MF_BYCOMMAND | (enable ? MF_CHECKED : MF_UNCHECKED));
    streamControls();
    timing();
    updateTaskbarProgress();
    invalidateButtons();
}
void command(int id)
{
    if (blocked())
        return;
    if (id == menu_mnuShuffle)
        id = main_chkShuffle;
    else if (id == menu_mnuRepeat)
        id = main_chkRepeat;
    if (id >= menu_streamBase && id < menu_streamBase + app->streams.count)
    {
        Text address = app->streams[id - menu_streamBase];
        loadFile(address);
        if (app->playback.hasMedia())
            persist(L"General", L"LastStreamURL", address);
        return;
    }
    switch (id)
    {
    case main_cmdPlay:
    case compact_cmdPlay:
        play();
        break;
    case main_cmdPause:
    case compact_cmdPause:
        pause();
        break;
    case main_cmdStop:
    case compact_cmdStop:
        stop();
        break;
    case main_cmdPrev:
    case compact_cmdPrev:
        nextTrack(true);
        break;
    case main_cmdnext:
    case compact_cmdNext:
        nextTrack();
        break;
    case main_chkRepeat:
    case main_chkShuffle:
        if (id == main_chkRepeat && app->isStream)
            break;
        app->checked[id] = !app->checked[id];
        CheckMenuItem(app->menu, id == main_chkRepeat ? menu_mnuRepeat : menu_mnuShuffle,
                      MF_BYCOMMAND | (app->checked[id] ? MF_CHECKED : MF_UNCHECKED));
        invalidateButtons();
        break;
    case menu_mnuReverse:
        toggleReverse();
        break;
    case main_cmdApply:
        app->checked[id] = !app->checked[id];
        applyRepeat();
        invalidateButtons();
        break;
    case compact_cmdVolDown:
        volume(app->volume - app->compactStep);
        break;
    case compact_cmdVolUp:
        volume(app->volume + app->compactStep);
        break;
    case compact_cmdReturnToMain:
        returnMain();
        break;
    case menu_AddStream:
        addStream();
        break;
    case menu_mnuClose:
        shutdown();
        break;
    case menu_mnuAbout:
        message(L"Stardisk's Player v1.2.1\r\nАвтор: Александров Олег Игоревич aka Stardisk", MB_ICONINFORMATION);
        break;
    case menu_mnuGoTo:
    {
        Text path = trim(inputBox(L"Введите путь к папке:", app->folder));
        if (!path.empty())
            setFolder(path);
        break;
    }
    case menu_mnuSearch:
    {
        Text text = trim(inputBox(L"Название файла или его часть:"));
        if (text.empty())
            break;
        bool found = false;
        Text needle = lower(text);
        for (int i = 0; i < fileCount(); ++i)
        {
            Text name = lower(fileName(i));
            for (int j = 0; j + needle.size() <= name.size(); ++j)
                if (equal(name.part(j, needle.size()), needle))
                {
                    selectTrack(i);
                    found = true;
                    break;
                }
            if (found)
                break;
        }
        if (!found)
            message(L"Не найдено", MB_ICONINFORMATION);
        break;
    }
    case menu_mnuOpenLocation:
    {
        Text file = selectedFile();
        if (!file.empty())
        {
            PIDLIST_ABSOLUTE pidl = ILCreateFromPathW(file);
            if (pidl)
            {
                SHOpenFolderAndSelectItems(pidl, 0, nullptr, 0);
                ILFree(pidl);
            }
        }
        break;
    }
    case menu_mnuRename:
    case menu_mnuMove:
    case menu_mnuDelete:
        fileOperation(id);
        break;
    case menu_mnuGotoCompact:
        showCompact();
        break;
    case menu_mnuRestoreCompact:
        persist(L"CompactPosition", L"X", L"0");
        persist(L"CompactPosition", L"Y", L"0");
        showCompact();
        message(L"Компактный режим находится в верхнем левом углу.", MB_ICONINFORMATION);
        break;
    case menu_mnuCompactVolume:
    {
        Text value = inputBox(L"Укажите, на сколько процентов будет изменяться уровень громкости при нажатии кнопок "
                              L"управления громкостью в компактном режиме:",
                              readINI(app->presets, L"General", L"CompactStepVolume"), title);
        if (value.empty())
            break;
        int n;
        if (!integer(value, n))
        {
            message(L"Введите число", MB_OK);
            break;
        }
        if (n < 1 || n > 100)
        {
            message(L"Число должно быть от 1 до 100", MB_OK);
            break;
        }
        persist(L"General", L"CompactStepVolume", number(n));
        break;
    }
    case menu_mnuCloseToCompact:
        toggleView(id, app->closeCompact, L"CloseToCompact");
        break;
    case menu_mnuMinimizeToCompact:
        toggleView(id, app->minimizeCompact, L"MinimizeToCompact");
        break;
    case menu_mnuShowTrackInCaption:
        toggleView(id, app->showCaption, L"ShowTrackInCaption");
        updateCaption();
        break;
    case menu_mnuShowTaskbarProgress:
        toggleView(id, app->showTaskbarProgress, L"ShowTaskbarProgress");
        if (app->showTaskbarProgress)
        {
            if (app->status == 1 && !app->isStream)
                app->realPosition = app->playback.position();
            updateTaskbarProgress();
        }
        else if (app->taskbar)
            app->taskbar->SetProgressState(app->main, TBPF_NOPROGRESS);
        break;
    }
}

void tick()
{
    if (blocked() || !app->playback.hasMedia())
        return;
    bool ended = false;
    HRESULT hr = app->playback.pollEnded(ended);
    if (FAILED(hr))
    {
        if (!app->isStream)
        {
            playbackFailure(hr, app->playback.reverseError());
            return;
        }
        clearMedia(false);
        changeStatus(0);
        setText(main_lblSysInfo, L"Поток прерван");
        message(Text(L"Воспроизведение потока прервано:\r\n") + windowsError(hr) + L"\r\nКод: " + hexCode(hr));
        return;
    }
    if (ended)
        app->ended = true;
    if (app->isStream)
    {
        if (app->ended)
        {
            clearMedia(false);
            changeStatus(0);
            setText(main_lblSysInfo, L"Поток завершён");
        }
        else if (app->status == 1)
        {
            app->realPosition += 1;
            timing();
        }
        return;
    }
    if (app->status != 1)
        return;
    if (!app->repeating && app->ended)
    {
        if (app->checked[main_chkRepeat])
        {
            seek(app->playback.isReverse() ? app->playback.duration() : 0);
            app->playback.run();
            app->ended = false;
        }
        else
            nextTrack();
        return;
    }
    app->realPosition = app->playback.position();
    if (app->repeating &&
        (app->playback.isReverse()
             ? (app->realPosition <= app->start || (app->end && app->realPosition > app->end) || app->ended)
             : (app->realPosition < app->start || (app->end && app->realPosition >= app->end) || app->ended)))
    {
        seek(app->playback.isReverse() ? (app->end ? app->end : app->playback.duration()) : app->start);
        if (app->ended)
            app->playback.run();
        app->ended = false;
    }
    if (!app->dragging)
    {
        SendMessageW(control(main_hsprogress), TBM_SETPOS, TRUE, (LPARAM)app->realPosition);
        timing();
    }
    updateTaskbarProgress();
}
void mediaCommand(int cmd, int source)
{
    if (blocked() || !IsWindowEnabled(app->main))
        return;
    ULONGLONG now = GetTickCount64();
    if (cmd == app->lastCommand && source != app->lastSource && now - app->lastTick < 100)
        return;
    app->lastCommand = cmd;
    app->lastSource = source;
    app->lastTick = now;
    // Defer playback work until the native key message has unwound.
    PostMessageW(app->main, mediaMessage, cmd, 0);
}

COLORREF color(LONG ole)
{
    return ole < 0 ? GetSysColor(ole & 0xFF) : COLORREF(ole);
}
void iconLabelBackground(HDC dc, int valueId)
{
    // A transparent VB label overlays another lightweight label. Win32 static
    // controls are separate windows, so returning a hollow brush alone leaves
    // stale pixels. Repaint the underlying glyph in the value's clipped DC.
    int iconId = valueId == main_lblVolume ? main_Label3 : main_Label4;
    const ControlDef* value = definition(valueId);
    const ControlDef* icon = definition(iconId);
    RECT client{};
    GetClientRect(control(valueId), &client);
    FillRect(dc, &client, GetSysColorBrush(COLOR_BTNFACE));
    int saved = SaveDC(dc);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, color(icon->fore));
    SelectObject(dc, app->fonts[iconId]);
    RECT bounds{icon->x - value->x, icon->y - value->y, icon->x - value->x + icon->w, icon->y - value->y + icon->h};
    wchar_t symbol[2] = {wchar_t(0xF000 + *icon->text), 0};
    DrawTextW(dc, symbol, 1, &bounds, DT_LEFT | DT_TOP | DT_NOPREFIX);
    RestoreDC(dc, saved);
}
void drawButton(DRAWITEMSTRUCT* item)
{
    int id = item->CtlID;
    const ControlDef* d = definition(id);
    if (!d)
        return;
    bool active = (id == main_cmdPlay || id == compact_cmdPlay)     ? app->status == 1
                  : (id == main_cmdPause || id == compact_cmdPause) ? app->status == 2
                  : (id == main_cmdStop || id == compact_cmdStop)   ? app->status == 0
                                                                    : app->checked[id];
    bool down = (item->itemState & ODS_SELECTED) || app->checked[id];
    bool disabled = item->itemState & ODS_DISABLED;
    RECT r = item->rcItem;
    FillRect(item->hDC, &r, GetSysColorBrush(COLOR_BTNFACE));
    DrawFrameControl(item->hDC, &r, DFC_BUTTON,
                     DFCS_BUTTONPUSH | (down ? DFCS_PUSHED : 0) | (disabled ? DFCS_INACTIVE : 0));
    RECT inside = r;
    InflateRect(&inside, -2, -2);
    HBRUSH brush = CreateSolidBrush(active ? RGB(192, 192, 192) : GetSysColor(COLOR_BTNFACE));
    FillRect(item->hDC, &inside, brush);
    DeleteObject(brush);
    if (down)
        OffsetRect(&inside, 1, 1);
    SetBkMode(item->hDC, TRANSPARENT);
    SetTextColor(item->hDC, disabled ? GetSysColor(COLOR_GRAYTEXT) : GetSysColor(COLOR_BTNTEXT));
    SelectObject(item->hDC, app->fonts[id]);
    Text caption(windowText(control(id)));
    if (equal(d->font, L"Webdings"))
    {
        wchar_t symbol[2] = {wchar_t(0xF000 + *caption.c_str()), 0};
        caption = symbol;
    }
    DrawTextW(item->hDC, caption, -1, &inside, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    if (item->itemState & ODS_FOCUS)
    {
        RECT focus = r;
        InflateRect(&focus, -3, -3);
        DrawFocusRect(item->hDC, &focus);
    }
}
void drawDirectory(DRAWITEMSTRUCT* item)
{
    if (item->itemID == UINT(-1))
        return;
    HWND list = control(main_Dir1);
    bool selected = item->itemState & ODS_SELECTED;
    FillRect(item->hDC, &item->rcItem, GetSysColorBrush(selected ? COLOR_HIGHLIGHT : COLOR_WINDOW));
    int meta = (int)SendMessageW(list, LB_GETITEMDATA, item->itemID, 0);
    int x = item->rcItem.left + 2 + (meta >> 1) * 16;
    DrawIconEx(item->hDC, x, item->rcItem.top, meta & 1 ? app->folderOpen : app->folderClosed, 16, 16, 0, nullptr,
               DI_NORMAL);
    Text name = leaf(app->directoryPaths[item->itemID]);
    if (name.empty())
        name = app->directoryPaths[item->itemID];
    RECT r = item->rcItem;
    r.left = x + 18;
    SetBkMode(item->hDC, TRANSPARENT);
    SetTextColor(item->hDC, GetSysColor(selected ? COLOR_HIGHLIGHTTEXT : COLOR_WINDOWTEXT));
    SelectObject(item->hDC, app->fonts[main_Dir1]);
    DrawTextW(item->hDC, name, -1, &r, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);
    if (item->itemState & ODS_FOCUS)
        DrawFocusRect(item->hDC, &item->rcItem);
}
void drawDrive(DRAWITEMSTRUCT* item)
{
    if (item->itemID == UINT(-1))
        return;
    bool selected = item->itemState & ODS_SELECTED;
    FillRect(item->hDC, &item->rcItem, GetSysColorBrush(selected ? COLOR_HIGHLIGHT : COLOR_WINDOW));
    wchar_t root[4] = {wchar_t(SendMessageW(control(main_Drive1), CB_GETITEMDATA, item->itemID, 0)), L':', L'\\', 0};
    SHFILEINFOW info{};
    SHGetFileInfoW(root, 0, &info, sizeof(info), SHGFI_ICON | SHGFI_SMALLICON);
    if (info.hIcon)
    {
        DrawIconEx(item->hDC, item->rcItem.left + 2, item->rcItem.top, info.hIcon, 16, 16, 0, nullptr, DI_NORMAL);
        DestroyIcon(info.hIcon);
    }
    wchar_t buf[512]{};
    SendMessageW(control(main_Drive1), CB_GETLBTEXT, item->itemID, (LPARAM)buf);
    RECT r = item->rcItem;
    r.left += 22;
    SetBkMode(item->hDC, TRANSPARENT);
    SetTextColor(item->hDC, GetSysColor(selected ? COLOR_HIGHLIGHTTEXT : COLOR_WINDOWTEXT));
    SelectObject(item->hDC, app->fonts[main_Drive1]);
    DrawTextW(item->hDC, buf, -1, &r, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);
    if (item->itemState & ODS_FOCUS)
        DrawFocusRect(item->hDC, &r);
}
RECT compactTitleInterior(HWND h)
{
    RECT bounds{};
    GetClientRect(h, &bounds);
    InflateRect(&bounds, -GetSystemMetrics(SM_CXEDGE), -GetSystemMetrics(SM_CYEDGE));
    return bounds;
}
void paintCompactTitle(HWND h, HDC dc)
{
    RECT r{};
    GetClientRect(h, &r);
    DrawEdge(dc, &r, EDGE_SUNKEN, BF_RECT | BF_ADJUST);
    HBRUSH bg = CreateSolidBrush(RGB(224, 224, 224));
    FillRect(dc, &r, bg);
    DeleteObject(bg);
    RECT progress = r;
    progress.right = progress.left + LONG((r.right - r.left) * playbackFraction(app->playback.duration()));
    HBRUSH shade = CreateSolidBrush(RGB(176, 176, 176));
    FillRect(dc, &progress, shade);
    DeleteObject(shade);
    SaveDC(dc);
    SelectObject(dc, app->compactTitleFont);
    SetTextColor(dc, RGB(0, 0, 0));
    SetBkMode(dc, TRANSPARENT);
    Text name = fileName(selectedIndex());
    ExtTextOutW(dc, r.left + app->scrollX, r.top, ETO_CLIPPED, &r, name, name.size(), nullptr);
    RestoreDC(dc, -1);
}
LRESULT CALLBACK compactTitleProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    if (msg == WM_PAINT)
    {
        PAINTSTRUCT ps{};
        HDC dc = BeginPaint(h, &ps);
        paintCompactTitle(h, dc);
        EndPaint(h, &ps);
        return 0;
    }
    if (msg == WM_LBUTTONDOWN)
    {
        ReleaseCapture();
        SendMessageW(app->compact, WM_NCLBUTTONDOWN, HTCAPTION, 0);
        return 0;
    }
    return DefSubclassProc(h, msg, w, l);
}
LRESULT CALLBACK dragLabelProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    if (msg == WM_LBUTTONDOWN)
    {
        ReleaseCapture();
        SendMessageW(app->compact, WM_NCLBUTTONDOWN, HTCAPTION, 0);
        return 0;
    }
    return DefSubclassProc(h, msg, w, l);
}
void paintDisabledScrollbar(HWND h, HDC dc)
{
    RECT track{};
    GetClientRect(h, &track);
    FillRect(dc, &track, GetSysColorBrush(COLOR_SCROLLBAR));
    int arrowWidth = clamp(GetSystemMetrics(SM_CXHSCROLL), 0, (track.right - track.left) / 2);
    RECT left = track, right = track;
    left.right = left.left + arrowWidth;
    right.left = right.right - arrowWidth;
    DrawFrameControl(dc, &left, DFC_SCROLL, DFCS_SCROLLLEFT | DFCS_INACTIVE);
    DrawFrameControl(dc, &right, DFC_SCROLL, DFCS_SCROLLRIGHT | DFCS_INACTIVE);
}
LRESULT CALLBACK scrollbarProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    // Keep the enabled control native. A disabled classic scrollbar must also
    // look inactive, rather than leaving its normal thumb and dark arrows.
    if (!IsWindowEnabled(h))
    {
        if (msg == WM_PAINT)
        {
            PAINTSTRUCT ps{};
            HDC dc = BeginPaint(h, &ps);
            paintDisabledScrollbar(h, dc);
            EndPaint(h, &ps);
            return 0;
        }
        if (msg == WM_PRINTCLIENT)
        {
            paintDisabledScrollbar(h, reinterpret_cast<HDC>(w));
            return 0;
        }
    }
    LRESULT result = DefSubclassProc(h, msg, w, l);
    if (msg == WM_ENABLE)
        InvalidateRect(h, nullptr, TRUE);
    return result;
}
LRESULT CALLBACK volumeIconProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    if (msg == WM_CTLCOLORSTATIC && reinterpret_cast<HWND>(l) == control(main_lblVolume))
        return SendMessageW(GetParent(h), msg, w, l);
    return DefSubclassProc(h, msg, w, l);
}
LRESULT CALLBACK progressProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    MSG key{};
    key.hwnd = h;
    key.message = msg;
    key.wParam = w;
    key.lParam = l;
    if (handlePositionKey(key))
        return 0;
    if (msg == WM_LBUTTONDOWN && !blocked())
        app->dragging = true;
    LRESULT result = DefSubclassProc(h, msg, w, l);
    if ((msg == WM_LBUTTONUP || msg == WM_KILLFOCUS) && app->dragging)
    {
        app->dragging = false;
        if (!blocked())
            seek(double(SendMessageW(h, TBM_GETPOS, 0, 0)));
    }
    return result;
}
LRESULT CALLBACK buttonProc(HWND h, UINT msg, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR)
{
    int id = GetDlgCtrlID(h);
    if ((msg == WM_RBUTTONUP || msg == WM_CONTEXTMENU) && (id == main_cmdPlay || id == compact_cmdPlay))
    {
        command(menu_mnuReverse);
        return 0;
    }
    if (msg == WM_LBUTTONDBLCLK && (id == compact_cmdVolDown || id == compact_cmdVolUp))
        msg = WM_LBUTTONDOWN;
    return DefSubclassProc(h, msg, w, l);
}
void createControls(HWND parentWindow, bool compact)
{
    for (const auto& d : controlDefs)
    {
        if (bool(d.compact) != compact)
            continue;
        const wchar_t* cls = L"STATIC";
        DWORD style = WS_CHILD | WS_VISIBLE, ex = 0;
        switch (d.kind)
        {
        case 0:
            style |= SS_NOPREFIX | (d.align == 1 ? SS_RIGHT : d.align == 2 ? SS_CENTER : SS_LEFT);
            if (d.id == main_Label3)
                ex |= WS_EX_TRANSPARENT;
            if (d.id == main_lblSysInfo)
            {
                style |= WS_BORDER | SS_EDITCONTROL;
            }
            break;
        case 1:
        case 2:
            cls = L"BUTTON";
            style |= BS_OWNERDRAW | WS_TABSTOP;
            break;
        case 3:
            cls = L"BUTTON";
            style |= BS_GROUPBOX;
            break;
        case 4:
            cls = L"EDIT";
            style |= ES_AUTOHSCROLL | WS_TABSTOP;
            ex = WS_EX_CLIENTEDGE;
            break;
        case 5:
            cls = L"SCROLLBAR";
            style |= SBS_HORZ | WS_TABSTOP;
            break;
        case 6:
            cls = TRACKBAR_CLASSW;
            style |= TBS_HORZ | TBS_NOTICKS | WS_TABSTOP;
            break;
        case 7:
            cls = L"COMBOBOX";
            style |= CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_VSCROLL | WS_TABSTOP;
            break;
        case 8:
            cls = L"LISTBOX";
            style |= LBS_NOTIFY | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_TABSTOP;
            ex = WS_EX_CLIENTEDGE;
            break;
        case 9:
            cls = L"LISTBOX";
            style |= LBS_NOTIFY | LBS_SORT | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_TABSTOP;
            ex = WS_EX_CLIENTEDGE;
            break;
        case 10:
            style |= SS_NOTIFY;
            break;
        }
        Text caption = d.text;
        if (equal(d.font, L"Webdings") && d.kind == 0)
        {
            wchar_t glyph[2] = {wchar_t(0xF000 + *d.text), 0};
            caption = glyph;
        }
        HWND h = CreateWindowExW(ex, cls, caption, style, d.x, d.y, d.w, d.kind == 7 ? 250 : d.h, parentWindow,
                                 (HMENU)(INT_PTR)d.id, app->instance, nullptr);
        app->controls[d.id] = h;
        app->fonts[d.id] = CreateFontW(-MulDiv(d.points100, 96, 7200), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                       equal(d.font, L"Webdings") ? SYMBOL_CHARSET : DEFAULT_CHARSET,
                                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, d.font);
        SendMessageW(h, WM_SETFONT, (WPARAM)app->fonts[d.id], TRUE);
        SetWindowTheme(h, L"", L"");
        if (d.id == main_Label3)
            SetWindowSubclass(h, volumeIconProc, 1, 0);
        if (d.kind == 1 || d.kind == 2)
            SetWindowSubclass(h, buttonProc, 1, 0);
        if (d.kind == 5)
            SetWindowSubclass(h, scrollbarProc, 1, 0);
        if (d.kind == 4)
            SendMessageW(h, EM_SETLIMITTEXT, 5, 0);
        if (d.kind == 8)
            SendMessageW(h, LB_SETITEMHEIGHT, 0, 16);
        if (d.kind == 7)
        {
            SendMessageW(h, CB_SETITEMHEIGHT, 0, 16);
            SendMessageW(h, CB_SETITEMHEIGHT, -1, 17);
        }
        if (d.kind == 6)
        {
            SendMessageW(h, TBM_SETRANGEMAX, TRUE, 32767);
            SendMessageW(h, TBM_SETLINESIZE, 0, 10);
            SendMessageW(h, TBM_SETPAGESIZE, 0, 20);
            SetWindowSubclass(h, progressProc, 1, 0);
        }
        if (d.id == compact_Picture1)
            SetWindowSubclass(h, compactTitleProc, 1, 0);
        if (compact && d.kind == 0)
            SetWindowSubclass(h, dragLabelProc, 1, 0);
    }
    // Preserve the designer's tab order despite its independent paint order.
    for (int tab = 30; tab >= 0; --tab)
        for (const auto& d : controlDefs)
            if (bool(d.compact) == compact && d.tab == tab)
                SetWindowPos(control(d.id), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    if (!compact)
    {
        // A child label is painted after its icon parent, including the very
        // first frame. Transparent sibling labels can paint in reverse order.
        const ControlDef* icon = definition(main_Label3);
        const ControlDef* value = definition(main_lblVolume);
        SetParent(control(main_lblVolume), control(main_Label3));
        SetWindowPos(control(main_lblVolume), HWND_TOP, value->x - icon->x, value->y - icon->y, value->w, value->h,
                     SWP_NOACTIVATE);
        SetScrollRange(control(main_hsvolume), SB_CTL, 0, 100, FALSE);
        SetScrollRange(control(main_hsRate), SB_CTL, 50, 200, FALSE);
    }
}
void createTooltips()
{
    app->tooltip =
        CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr, WS_POPUP | TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT,
                        CW_USEDEFAULT, CW_USEDEFAULT, app->main, nullptr, app->instance, nullptr);
    SendMessageW(app->tooltip, TTM_SETMAXTIPWIDTH, 0, 600);
    for (const auto& d : controlDefs)
    {
        const wchar_t* tip = d.tip;
        if ((d.id == main_cmdPlay || d.id == compact_cmdPlay) && !app->playback.reverseAvailable())
            tip = L"Воспроизведение";
        if (d.id == main_hsprogress && !*tip)
            tip = L"Позиция воспроизведения";
        if (d.id == main_Label4)
            tip = L"Скорость воспроизведения";
        if (!*tip && d.id != main_lblSysInfo)
            continue;
        TOOLINFOW tool{};
        tool.cbSize = TTTOOLINFOW_V2_SIZE;
        tool.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        tool.hwnd = d.compact ? app->compact : app->main;
        tool.uId = (UINT_PTR)control(d.id);
        tool.lpszText = d.id == main_lblSysInfo ? LPSTR_TEXTCALLBACKW : const_cast<wchar_t*>(tip);
        SendMessageW(app->tooltip, TTM_ADDTOOLW, 0, (LPARAM)&tool);
    }
}
HMENU submenu(const MenuItem* items, int count)
{
    HMENU menu = CreatePopupMenu();
    bool reverseAvailable = app->playback.reverseAvailable();
    for (int i = 0; i < count; ++i)
    {
        if (!reverseAvailable &&
            (items[i].id == menu_mnuReverse || (!items[i].id && i + 1 < count && items[i + 1].id == menu_mnuReverse)))
            continue;
        AppendMenuW(menu, items[i].id ? MF_STRING : MF_SEPARATOR, items[i].id, items[i].id ? items[i].text : nullptr);
    }
    return menu;
}
void createMenu()
{
    app->menu = CreateMenu();
    for (const auto& d : menuDefs)
    {
        HMENU popup = submenu(d.items, d.count);
        if (d.streams)
            app->streamsMenu = popup;
        AppendMenuW(app->menu, MF_POPUP, (UINT_PTR)popup, d.caption);
    }
}
LRESULT CALLBACK windowProc(HWND h, UINT msg, WPARAM w, LPARAM l)
{
    bool compact = h == app->compact;
    if (h == app->main && app->taskbarButtonCreated && msg == app->taskbarButtonCreated)
    {
        if (!app->taskbar)
        {
            ITaskbarList3* taskbar = nullptr;
            if (SUCCEEDED(CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&taskbar))))
            {
                if (SUCCEEDED(taskbar->HrInit()))
                    app->taskbar = taskbar;
                else
                    taskbar->Release();
            }
        }
        updateTaskbarProgress();
        return 0;
    }
    switch (msg)
    {
    case WM_DRAWITEM:
    {
        auto item = reinterpret_cast<DRAWITEMSTRUCT*>(l);
        if (item->CtlID == main_Dir1)
            drawDirectory(item);
        else if (item->CtlID == main_Drive1)
            drawDrive(item);
        else
            drawButton(item);
        return TRUE;
    }
    case WM_MEASUREITEM:
    {
        auto item = reinterpret_cast<MEASUREITEMSTRUCT*>(l);
        item->itemHeight = 16;
        return TRUE;
    }
    case WM_CTLCOLORSTATIC:
    {
        HDC dc = (HDC)w;
        int id = GetDlgCtrlID((HWND)l);
        const ControlDef* d = definition(id);
        bool overlay = id == main_lblVolume || id == main_lblRate;
        if (overlay)
            iconLabelBackground(dc, id);
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, d ? color(d->fore) : GetSysColor(COLOR_BTNTEXT));
        return (LRESULT)((overlay || id == main_Label3) ? GetStockObject(NULL_BRUSH) : GetSysColorBrush(COLOR_BTNFACE));
    }
    case WM_COMMAND:
    {
        int id = LOWORD(w), notice = HIWORD(w);
        if (blocked())
            return 0;
        if (id == main_File1 && notice == LBN_SELCHANGE)
        {
            loadFile(selectedFile());
            return 0;
        }
        if (id == main_Dir1 && notice == LBN_DBLCLK)
        {
            int index = (int)SendMessageW(control(id), LB_GETCURSEL, 0, 0);
            if (index >= 0 && index < app->directoryPaths.count)
            {
                Text path = app->directoryPaths[index];
                setFolder(path);
            }
            return 0;
        }
        if (id == main_Drive1 && notice == CBN_SELCHANGE)
        {
            int i = (int)SendMessageW(control(id), CB_GETCURSEL, 0, 0);
            wchar_t root[4] = {wchar_t(SendMessageW(control(id), CB_GETITEMDATA, i, 0)), L':', L'\\', 0};
            if (!setFolder(Text(root)))
                setFolder(app->home);
            return 0;
        }
        if (!l || notice == BN_CLICKED)
            command(id);
        return 0;
    }
    case WM_HSCROLL:
    {
        if (blocked())
            return 0;
        HWND bar = (HWND)l;
        int id = GetDlgCtrlID(bar), code = LOWORD(w);
        if (id == main_hsprogress)
        {
            int value = (int)SendMessageW(bar, TBM_GETPOS, 0, 0);
            if (app->dragging)
                setText(main_lblTiming, timeText(value) + L" / " + timeText(int(app->playback.duration())));
            else if (code != TB_ENDTRACK)
                seek(value);
            return 0;
        }
        if (id == main_hsvolume || id == main_hsRate)
        {
            int value = id == main_hsvolume ? app->volume : app->speed;
            switch (code)
            {
            case SB_LINELEFT:
                --value;
                break;
            case SB_LINERIGHT:
                ++value;
                break;
            case SB_PAGELEFT:
                value -= 10;
                break;
            case SB_PAGERIGHT:
                value += 10;
                break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION:
                value = HIWORD(w);
                break;
            case SB_LEFT:
                value = id == main_hsvolume ? 0 : 50;
                break;
            case SB_RIGHT:
                value = id == main_hsvolume ? 100 : 200;
                break;
            default:
                return 0;
            }
            if (id == main_hsvolume)
                volume(value);
            else
                speed(value);
        }
        return 0;
    }
    case WM_NOTIFY:
    {
        auto n = (NMHDR*)l;
        if (n->hwndFrom == app->tooltip && n->code == TTN_GETDISPINFOW)
        {
            auto info = (NMTTDISPINFOW*)l;
            info->lpszText = const_cast<wchar_t*>(app->info.c_str());
            return 0;
        }
        break;
    }
    case WM_TIMER:
        if (w == 1)
            tick();
        else if (w == 2)
        {
            setText(compact_Label2, windowText(control(main_lblTiming)));
            HDC dc = GetDC(control(compact_Picture1));
            SelectObject(dc, app->compactTitleFont);
            Text name = fileName(selectedIndex());
            SIZE size{};
            GetTextExtentPoint32W(dc, name, name.size(), &size);
            ReleaseDC(control(compact_Picture1), dc);
            app->scrollWidth = size.cx;
            app->scrollX -= 5;
            if (app->scrollX < -app->scrollWidth)
            {
                RECT bounds = compactTitleInterior(control(compact_Picture1));
                app->scrollX = bounds.right - bounds.left;
            }
            InvalidateRect(control(compact_Picture1), nullptr, FALSE);
        }
        return 0;
    case WM_HOTKEY:
        if (w >= 0x5110 && w < 0x5114)
        {
            const int commands[] = {14, 13, 12, 11};
            mediaCommand(commands[w - 0x5110], 1);
        }
        return 0;
    case WM_APPCOMMAND:
    {
        int cmd = GET_APPCOMMAND_LPARAM(l);
        if (cmd == 11 || cmd == 12 || cmd == 13 || cmd == 14 || cmd == 46 || cmd == 47)
        {
            mediaCommand(cmd, 2);
            return TRUE;
        }
        break;
    }
    case mediaMessage:
        if (!blocked())
        {
            switch (w)
            {
            case 14:
                if (app->status == 1)
                    pause();
                else
                    play();
                break;
            case 13:
                stop();
                break;
            case 12:
                nextTrack(true);
                break;
            case 11:
                nextTrack();
                break;
            case 46:
                if (app->status != 1)
                    play();
                break;
            case 47:
                if (app->status == 1)
                    pause();
                break;
            }
        }
        return 0;
    case startupMessage:
        if (fileCount())
            selectTrack(selectedIndex() < 0 ? 0 : selectedIndex());
        return 0;
    case WM_SIZE:
        if (!compact && w == SIZE_MINIMIZED && app->minimizeCompact && !blocked())
            showCompact();
        return 0;
    case WM_LBUTTONDOWN:
        if (compact)
        {
            ReleaseCapture();
            SendMessageW(h, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            return 0;
        }
        break;
    case WM_CLOSE:
        if (blocked())
            return 0;
        if (!compact && app->closeCompact)
            showCompact();
        else
            shutdown();
        return 0;
    case WM_DESTROY:
        if (!compact)
            PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(h, msg, w, l);
}
void initialize()
{
    auto path = static_cast<wchar_t*>(allocate(65536));
    GetModuleFileNameW(nullptr, path, 32768);
    app->home = parent(Text(path));
    freeMemory(path);
    app->presets = join(app->home, L"presets.ini");
    app->repeatlist = join(app->home, L"repeatlist.ini");
    if (!exists(app->presets))
    {
        ensureINI(app->presets);
        persist(L"General", L"LastPath", app->home);
        persist(L"General", L"CompactStepVolume", L"5");
    }
    app->random = GetTickCount() | 1;
    WNDCLASSEXW cls{};
    cls.cbSize = sizeof(cls);
    cls.style = CS_HREDRAW | CS_VREDRAW;
    cls.lpfnWndProc = windowProc;
    cls.hInstance = app->instance;
    cls.hIcon = LoadIconW(app->instance, MAKEINTRESOURCEW(1));
    cls.hIconSm =
        static_cast<HICON>(LoadImageW(app->instance, MAKEINTRESOURCEW(1), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON),
                                      GetSystemMetrics(SM_CYSMICON), LR_SHARED));
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.hbrBackground = GetSysColorBrush(COLOR_BTNFACE);
    cls.lpszClassName = L"StardiskPlayerCpp";
    RegisterClassExW(&cls);
    app->taskbarButtonCreated = RegisterWindowMessageW(L"TaskbarButtonCreated");
    createMenu();
    RECT r{0, 0, 353, 462};
    AdjustWindowRectEx(&r, mainStyle, TRUE, 0);
    app->main = CreateWindowExW(0, cls.lpszClassName, title, mainStyle, CW_USEDEFAULT, CW_USEDEFAULT, r.right - r.left,
                                r.bottom - r.top, nullptr, app->menu, app->instance, nullptr);
    app->compact =
        CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, cls.lpszClassName, L"Stardisk's Player Compact Mode",
                        WS_POPUP, 0, 0, 301, 17, nullptr, nullptr, app->instance, nullptr);
    createControls(app->main, false);
    createControls(app->compact, true);
    app->compactTitleFont =
        CreateFontW(-9, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                    CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, L"Tahoma");
    SHSTOCKICONINFO icon{};
    icon.cbSize = sizeof(icon);
    if (SUCCEEDED(SHGetStockIconInfo(SIID_FOLDER, SHGSI_ICON | SHGSI_SMALLICON, &icon)))
        app->folderClosed = icon.hIcon;
    if (SUCCEEDED(SHGetStockIconInfo(SIID_FOLDEROPEN, SHGSI_ICON | SHGSI_SMALLICON, &icon)))
        app->folderOpen = icon.hIcon;
    DWORD drives = GetLogicalDrives();
    for (int i = 0; i < 26; ++i)
        if (drives & (1u << i))
        {
            wchar_t root[4] = {wchar_t(L'A' + i), L':', L'\\', 0}, label[261]{};
            GetVolumeInformationW(root, label, 261, nullptr, nullptr, nullptr, nullptr, 0);
            Text caption = Text(root).part(0, 2);
            if (*label)
                caption += Text(L" [") + label + L"]";
            int index = (int)SendMessageW(control(main_Drive1), CB_ADDSTRING, 0, (LPARAM)caption.c_str());
            SendMessageW(control(main_Drive1), CB_SETITEMDATA, index, L'A' + i);
        }
    Text folder = readINI(app->presets, L"General", L"LastPath");
    if (!exists(folder, true))
        folder = app->home;
    setFolder(folder, false);
    selectTrack(setting(L"General", L"LastIndex", 0));
    loadStreams();
    volume(clamp(setting(L"General", L"Volume", 100), 0, 100));
    speed(100);
    setText(main_lblSysInfo, L"");
    app->showCaption = setting(L"View", L"ShowTrackInCaption", 0) == 1;
    app->showTaskbarProgress = setting(L"View", L"ShowTaskbarProgress", 0) == 1;
    CheckMenuItem(app->menu, menu_mnuShowTaskbarProgress,
                  MF_BYCOMMAND | (app->showTaskbarProgress ? MF_CHECKED : MF_UNCHECKED));
    app->minimizeCompact = setting(L"View", L"MinimizeToCompact", 0) == 1;
    app->closeCompact = setting(L"View", L"CloseToCompact", 0) == 1;
    CheckMenuItem(app->menu, menu_mnuShowTrackInCaption, MF_BYCOMMAND | (app->showCaption ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(app->menu, menu_mnuMinimizeToCompact,
                  MF_BYCOMMAND | (app->minimizeCompact ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(app->menu, menu_mnuCloseToCompact, MF_BYCOMMAND | (app->closeCompact ? MF_CHECKED : MF_UNCHECKED));
    createTooltips();
    app->initializing = false;
    ShowWindow(app->main, SW_SHOWNORMAL);
    UpdateWindow(app->main);
    Text warning;
    const UINT keys[] = {VK_MEDIA_PLAY_PAUSE, VK_MEDIA_STOP, VK_MEDIA_PREV_TRACK, VK_MEDIA_NEXT_TRACK};
    const wchar_t* labels[] = {L"Play/Pause", L"Stop", L"Previous", L"Next"};
    for (int i = 0; i < 4; ++i)
    {
        app->registered[i] = !!RegisterHotKey(app->main, 0x5110 + i, MOD_NOREPEAT, keys[i]);
        if (!app->registered[i])
        {
            if (!warning.empty())
                warning += L", ";
            warning += labels[i];
        }
    }
    if (!warning.empty())
        message(Text(L"Не удалось зарегистрировать мультимедийные кнопки: ") + warning +
                L". Возможно, они заняты другим приложением.");
    PostMessageW(app->main, startupMessage, 0, 0);
}
} // namespace

#ifndef PLAYER_TEST
extern "C" void playerEntry()
{
    app = new (allocate(sizeof(App))) App();
    app->instance = GetModuleHandleW(nullptr);
    app->mutex = CreateMutexW(nullptr, FALSE, L"Local\\StardiskPlayerCpp-v2");
    if (GetLastError() == ERROR_ALREADY_EXISTS)
        ExitProcess(0);
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        message(windowsError(hr));
        ExitProcess(1);
    }
    INITCOMMONCONTROLSEX common{sizeof(common), ICC_BAR_CLASSES | ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&common);
    initialize();
    ACCEL shortcuts[] = {{FVIRTKEY | FCONTROL, 'L', menu_AddStream},
                         {FVIRTKEY | FCONTROL, 'F', menu_mnuSearch},
                         {FVIRTKEY | FCONTROL, 'E', menu_mnuOpenLocation},
                         {FVIRTKEY | FCONTROL, 'X', menu_mnuMove},
                         {FVIRTKEY, VK_DELETE, menu_mnuDelete},
                         {FVIRTKEY, VK_F2, menu_mnuRename}};
    HACCEL accel = CreateAcceleratorTableW(shortcuts, sizeof(shortcuts) / sizeof(shortcuts[0]));
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        HWND active = IsWindowVisible(app->main) ? app->main : app->compact;
        bool editing = GetFocus() == control(main_txtStartTime) || GetFocus() == control(main_txtEndTime);
        if (handlePositionKey(msg))
            continue;
        if ((!editing && TranslateAcceleratorW(app->main, accel, &msg)) || IsDialogMessageW(active, &msg))
            continue;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    DestroyAcceleratorTable(accel);
    if (app->taskbar)
        app->taskbar->Release();
    if (app->folderClosed)
        DestroyIcon(app->folderClosed);
    if (app->folderOpen)
        DestroyIcon(app->folderOpen);
    for (auto font : app->fonts)
        if (font)
            DeleteObject(font);
    DeleteObject(app->compactTitleFont);
    if (app->mutex)
        CloseHandle(app->mutex);
    app->~App();
    freeMemory(app);
    CoUninitialize();
    ExitProcess(0);
}
#endif
