// In-process integration tests of the application model and its own native
// controls. No automation of another running player or the user's windows.
#define PLAYER_TEST
#include "../src/player.cpp"

namespace
{
HANDLE testLog;
int failures = 0;
void log(const Text& text)
{
    Text line = text + L"\r\n";
    int count = WideCharToMultiByte(CP_UTF8, 0, line, line.size(), nullptr, 0, nullptr, nullptr);
    auto bytes = static_cast<char*>(allocate(count));
    WideCharToMultiByte(CP_UTF8, 0, line, line.size(), bytes, count, nullptr, nullptr);
    DWORD n;
    WriteFile(testLog, bytes, count, &n, nullptr);
    freeMemory(bytes);
}
void check(bool ok, const wchar_t* label)
{
    log(Text(ok ? L"PASS " : L"FAIL ") + label);
    if (!ok)
        ++failures;
}
Text menuCaption(HMENU menu, int row)
{
    wchar_t text[1024]{};
    GetMenuStringW(menu, row, text, 1024, MF_BYPOSITION);
    return text;
}
int paintedTextRows(HWND label)
{
    RECT rect{};
    GetClientRect(label, &rect);
    HDC dc = CreateCompatibleDC(nullptr);
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = rect.right;
    info.bmiHeader.biHeight = -rect.bottom;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ oldBitmap = SelectObject(dc, bitmap);
    FillRect(dc, &rect, GetSysColorBrush(COLOR_BTNFACE));
    SendMessageW(label, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT);
    SelectObject(dc, reinterpret_cast<HFONT>(SendMessageW(label, WM_GETFONT, 0, 0)));
    TEXTMETRICW metrics{};
    GetTextMetricsW(dc, &metrics);
    int rows = 0;
    for (int top = 0; top + metrics.tmHeight <= rect.bottom; top += metrics.tmHeight)
    {
        bool painted = false;
        for (int y = top + 1; y < top + metrics.tmHeight && !painted; ++y)
            for (int x = 2; x < rect.right - 2; ++x)
                if (GetPixel(dc, x, y) == GetSysColor(COLOR_BTNTEXT))
                {
                    painted = true;
                    break;
                }
        if (painted)
            ++rows;
    }
    SelectObject(dc, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(dc);
    return rows;
}
int paintedVolumeNumberPixels()
{
    RECT rect{};
    GetClientRect(control(main_Label3), &rect);
    HDC dc = CreateCompatibleDC(nullptr);
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = rect.right;
    info.bmiHeader.biHeight = -rect.bottom;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ old = SelectObject(dc, bitmap);
    FillRect(dc, &rect, GetSysColorBrush(COLOR_BTNFACE));
    SendMessageW(control(main_Label3), WM_PRINT, reinterpret_cast<WPARAM>(dc),
                 PRF_CLIENT | PRF_CHILDREN | PRF_ERASEBKGND);
    const ControlDef* icon = definition(main_Label3);
    const ControlDef* value = definition(main_lblVolume);
    int visible = 0;
    for (int y = value->y - icon->y; y < value->y - icon->y + value->h; ++y)
        for (int x = value->x - icon->x; x < value->x - icon->x + value->w; ++x)
            if (GetPixel(dc, x, y) == color(value->fore))
                ++visible;
    SelectObject(dc, old);
    DeleteObject(bitmap);
    DeleteDC(dc);
    return visible;
}
bool arrowKey(HWND target, UINT key, UINT message = WM_KEYDOWN)
{
    MSG msg{};
    msg.hwnd = target;
    msg.message = message;
    msg.wParam = key;
    return handlePositionKey(msg);
}
} // namespace

extern "C" void logicTestEntry()
{
    auto p = static_cast<wchar_t*>(allocate(65536));
    GetModuleFileNameW(nullptr, p, 32768);
    Text root = parent(parent(Text(p)));
    freeMemory(p);
    Text qa = join(root, L"qa"), logic = join(qa, L"logic");
    testLog = CreateFileW(join(qa, L"player-tests.log"), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                          FILE_ATTRIBUTE_NORMAL, nullptr);
    if (testLog == INVALID_HANDLE_VALUE)
        ExitProcess(100);
    HRESULT initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(initialized))
    {
        log(windowsError(initialized));
        ExitProcess(101);
    }
    app = new (allocate(sizeof(App))) App();
    app->instance = GetModuleHandleW(nullptr);
    app->home = logic;
    app->presets = join(logic, L"presets.ini");
    app->repeatlist = join(logic, L"repeatlist.ini");
    app->volume = 0;
    app->random = 1234567;
    INITCOMMONCONTROLSEX common{sizeof(common), ICC_BAR_CLASSES | ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&common);
    WNDCLASSEXW cls{};
    cls.cbSize = sizeof(cls);
    cls.lpfnWndProc = windowProc;
    cls.hInstance = app->instance;
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.hbrBackground = GetSysColorBrush(COLOR_BTNFACE);
    cls.lpszClassName = L"StardiskPlayerLogicTests";
    RegisterClassExW(&cls);
    createMenu();
    app->main = CreateWindowExW(0, cls.lpszClassName, title, mainStyle, 0, 0, 353, 462, nullptr, app->menu,
                                app->instance, nullptr);
    app->compact = CreateWindowExW(WS_EX_TOOLWINDOW, cls.lpszClassName, title, WS_POPUP, 0, 0, 301, 17, nullptr,
                                   nullptr, app->instance, nullptr);
    createControls(app->main, false);
    createControls(app->compact, true);
    volume(56);
    check(equal(windowText(control(main_lblVolume)), L"56") && paintedVolumeNumberPixels() > 0,
          L"Startup volume number paints white over speaker without user interaction");
    volume(0);
    createTooltips();
    log(Text(L"Tooltip tool count: ") + number(SendMessageW(app->tooltip, TTM_GETTOOLCOUNT, 0, 0)));
    check(equal(menuCaption(app->menu, 0), L"Файл") && equal(menuCaption(app->menu, 1), L"Режим") &&
              equal(menuCaption(app->menu, 2), L"Вид") && equal(menuCaption(app->menu, 3), L"Потоки"),
          L"Exact top menu order");
    HMENU file = GetSubMenu(app->menu, 0), view = GetSubMenu(app->menu, 2), mode = GetSubMenu(app->menu, 1);
    bool hasReverse = app->playback.reverseAvailable();
    check(GetMenuItemCount(mode) == (hasReverse ? 4 : 2) && GetMenuItemID(mode, 0) == menu_mnuShuffle &&
              GetMenuItemID(mode, 1) == menu_mnuRepeat && (!hasReverse || GetMenuItemID(mode, 3) == menu_mnuReverse),
          L"Mode menu hides reverse and separator when adjacent DLL is absent");
    for (int id : {main_cmdPlay, main_cmdPause, main_cmdStop, main_chkRepeat, main_chkShuffle, compact_cmdPlay,
                   compact_cmdVolDown, compact_cmdVolUp})
    {
        TOOLINFOW tip{};
        wchar_t tipText[1024]{};
        tip.cbSize = TTTOOLINFOW_V2_SIZE;
        tip.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        tip.hwnd = definition(id)->compact ? app->compact : app->main;
        tip.uId = (UINT_PTR)control(id);
        tip.lpszText = tipText;
        check(SendMessageW(app->tooltip, TTM_GETTOOLINFOW, 0, (LPARAM)&tip) && tip.lpszText && *tip.lpszText,
              L"Transport and mode controls have native tooltips");
    }
    const ControlDef* volumeIcon = definition(main_Label3);
    check(volumeIcon->y == 333 && volumeIcon->h == 39 && volumeIcon->points100 == 2775,
          L"Volume icon original position, size and font restored");
    HDC iconDC = GetDC(control(main_Label3));
    check((GetWindowLongPtrW(control(main_Label3), GWL_EXSTYLE) & WS_EX_TRANSPARENT) &&
              SendMessageW(app->main, WM_CTLCOLORSTATIC, (WPARAM)iconDC, (LPARAM)control(main_Label3)) ==
                  (LRESULT)GetStockObject(NULL_BRUSH),
          L"Volume icon paints with a transparent label background");
    ReleaseDC(control(main_Label3), iconDC);
    check(GetMenuItemID(file, 0) == menu_mnuRename && GetMenuItemID(file, 1) == menu_mnuMove &&
              GetMenuItemID(file, 2) == menu_mnuDelete && GetMenuItemID(file, 9) == menu_mnuClose,
          L"Exact File menu order");
    check(GetMenuItemID(view, 0) == menu_mnuShowTrackInCaption && GetMenuItemID(view, 6) == menu_mnuRestoreCompact,
          L"Exact View menu order");
    refreshStreams();
    check(GetMenuItemCount(app->streamsMenu) == 2 && GetMenuItemID(app->streamsMenu, 1) == menu_AddStream,
          L"Empty stream menu order");
    app->streams.add(Text(L"http://127.0.0.1:52751/test.wav?a=1&b=2"));
    saveStreams();
    refreshStreams();
    check(GetMenuItemCount(app->streamsMenu) == 3 && GetMenuItemID(app->streamsMenu, 0) == menu_streamBase &&
              GetMenuItemID(app->streamsMenu, 2) == menu_AddStream,
          L"Saved stream menu order");
    check(equal(menuCaption(app->streamsMenu, 0), L"http://127.0.0.1:52751/test.wav?a=1&&b=2"),
          L"Ampersands do not become URL mnemonics");
    app->streams.clear();
    loadStreams();
    check(app->streams.count == 1 && equal(app->streams[0], L"http://127.0.0.1:52751/test.wav?a=1&b=2"),
          L"Reload saved streams");
    check(setFolder(join(qa, L"audio"), false) && fileCount() == 2, L"Folder filtering and file enumeration");
    selectTrack(0);
    app->initializing = false;
    selectTrack(0);
    check(app->status == 1 && selectedIndex() == 0 && app->playback.hasMedia(),
          L"Selecting a local file starts playback");
    check(equal(leaf(app->current), L"01 - ÄÖÜ тест.wav"), L"Windows file sort and Unicode track selection");
    command(main_cmdPause);
    seek(15);
    int originalIndex = selectedIndex();
    check(arrowKey(control(main_Drive1), VK_LEFT) && app->realPosition > 4.99 && app->realPosition < 5.01,
          L"Left arrow on drive control seeks backward ten seconds");
    check(arrowKey(control(main_Drive1), VK_LEFT, WM_KEYUP) && app->realPosition > 4.99 && app->realPosition < 5.01,
          L"Arrow release is consumed without a second seek");
    check(arrowKey(control(main_File1), VK_RIGHT) && app->realPosition > 14.99 && app->realPosition < 15.01 &&
              selectedIndex() == originalIndex && app->status == 2,
          L"Right arrow seeks forward ten seconds without changing selection or pause");
    check(!arrowKey(control(main_txtStartTime), VK_LEFT) && app->realPosition > 14.99,
          L"Time edit fields keep normal caret arrow keys");
    seek(2);
    arrowKey(control(main_cmdPlay), VK_LEFT);
    check(app->realPosition == 0, L"Arrow seek clamps to beginning");
    seek(app->playback.duration() - 2);
    arrowKey(control(main_cmdPlay), VK_RIGHT);
    check(app->realPosition == app->playback.duration(), L"Arrow seek clamps to end");
    seek(15);
    SendMessageW(control(main_hsprogress), WM_KEYDOWN, VK_LEFT, 0);
    SendMessageW(control(main_hsprogress), WM_KEYUP, VK_LEFT, 0);
    check(app->realPosition > 4.99 && app->realPosition < 5.01,
          L"Focused progress control also seeks by ten rather than one second");
    SendMessageW(control(main_hsprogress), WM_KEYDOWN, VK_RIGHT, 0);
    check(app->realPosition > 14.99 && app->realPosition < 15.01,
          L"Focused progress right arrow seeks forward ten seconds");
    seek(0);
    command(main_cmdPlay);
    command(menu_mnuShuffle);
    check(app->checked[main_chkShuffle] && (GetMenuState(app->menu, menu_mnuShuffle, MF_BYCOMMAND) & MF_CHECKED),
          L"Shuffle menu presses button and checks menu");
    command(main_chkShuffle);
    check(!app->checked[main_chkShuffle] && !(GetMenuState(app->menu, menu_mnuShuffle, MF_BYCOMMAND) & MF_CHECKED),
          L"Shuffle button clears matching menu check");
    command(menu_mnuRepeat);
    check(app->checked[main_chkRepeat] && (GetMenuState(app->menu, menu_mnuRepeat, MF_BYCOMMAND) & MF_CHECKED),
          L"Repeat menu presses button and checks menu");
    command(main_chkRepeat);
    check(!app->checked[main_chkRepeat] && !(GetMenuState(app->menu, menu_mnuRepeat, MF_BYCOMMAND) & MF_CHECKED),
          L"Repeat button clears matching menu check");
    speed(150);
    if (hasReverse)
    {
        seek(10);
        command(main_cmdPause);
        command(menu_mnuReverse);
        check(app->reverseEnabled && app->playback.isReverse() && equal(windowText(control(main_cmdPlay)), L"3") &&
                  equal(windowText(control(compact_cmdPlay)), L"3") &&
                  (GetMenuState(app->menu, menu_mnuReverse, MF_BYCOMMAND) & MF_CHECKED),
              L"Reverse DLL synchronizes menu and both play glyphs");
        check(app->speed == 100 && app->status == 2 && !IsWindowEnabled(control(main_hsRate)) &&
                  app->realPosition > 9.9 && app->realPosition < 10.1,
              L"Switch to reverse preserves paused position and disables speed at x1");
        arrowKey(control(main_Drive1), VK_RIGHT);
        check(app->realPosition > 19.99 && app->realPosition < 20.01, L"Reverse right arrow still adds ten seconds");
        arrowKey(control(main_Drive1), VK_LEFT);
        check(app->realPosition > 9.99 && app->realPosition < 10.01, L"Reverse left arrow still subtracts ten seconds");
        command(main_cmdPlay);
        Sleep(170);
        check(app->playback.position() < 9.95, L"Reverse DLL actually plays backward through player");
        command(main_cmdPause);
        SendMessageW(control(compact_cmdPlay), WM_CONTEXTMENU, (WPARAM)control(compact_cmdPlay), -1);
        check(!app->reverseEnabled && !app->playback.isReverse() && app->speed == 150 && app->status == 2 &&
                  equal(windowText(control(main_cmdPlay)), L"4") &&
                  !(GetMenuState(app->menu, menu_mnuReverse, MF_BYCOMMAND) & MF_CHECKED),
              L"Compact play switches back, restores vinyl speed and preserves pause");
        SendMessageW(control(main_cmdPlay), WM_RBUTTONUP, 0, MAKELPARAM(5, 5));
        check(app->reverseEnabled, L"Main play right click switches to reverse DLL");
        command(main_cmdPlay);
        setText(main_txtStartTime, L"00:02");
        setText(main_txtEndTime, L"00:04");
        command(main_cmdApply);
        check(app->repeating && app->realPosition > 3.9 && app->realPosition <= 4,
              L"Reverse segment repeat starts from B");
        seek(1.9);
        tick();
        check(app->realPosition > 3.9 && app->realPosition <= 4, L"Reverse segment boundary returns from A to B");
        command(main_cmdApply);
        command(main_cmdStop);
        check(app->status == 0 && app->realPosition == app->playback.duration(),
              L"Reverse stop resets position to track end");
        command(main_cmdnext);
        check(app->playback.isReverse() && app->realPosition > app->playback.duration() - .1,
              L"Next track retains reverse mode and starts at end");
        command(menu_mnuRepeat);
        seek(0);
        app->ended = true;
        tick();
        check(!app->ended && app->realPosition == app->playback.duration(),
              L"Whole-track reverse repeat returns to end");
        command(menu_mnuRepeat);
        loadFile(Text(L"http://127.0.0.1:52751/test.wav"));
        check(app->reverseEnabled && !app->playback.isReverse() && equal(windowText(control(main_cmdPlay)), L"4") &&
                  (GetMenuState(app->menu, menu_mnuReverse, MF_BYCOMMAND) & MF_GRAYED),
              L"Stream plays forward while retaining disabled reverse preference");
        command(menu_mnuReverse);
        check(app->reverseEnabled, L"Stream right-click/menu cannot change reverse preference");
        selectTrack(0);
        check(app->playback.isReverse() && equal(windowText(control(main_cmdPlay)), L"3"),
              L"Returning to local track restores selected reverse mode");
        command(menu_mnuReverse);
        seek(3);
        command(main_cmdPlay);
    }
    else
    {
        command(menu_mnuReverse);
        SendMessageW(control(main_cmdPlay), WM_RBUTTONUP, 0, MAKELPARAM(5, 5));
        SendMessageW(control(compact_cmdPlay), WM_CONTEXTMENU, (WPARAM)control(compact_cmdPlay), -1);
        check(!app->reverseEnabled && app->speed == 150 && equal(windowText(control(main_cmdPlay)), L"4"),
              L"Without DLL reverse menu command and both right clicks have no effect");
    }
    speed(100);
    volume(50);
    HWND volumeUp = control(compact_cmdVolUp);
    SendMessageW(volumeUp, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(5, 5));
    SendMessageW(volumeUp, WM_LBUTTONUP, 0, MAKELPARAM(5, 5));
    SendMessageW(volumeUp, WM_LBUTTONDBLCLK, MK_LBUTTON, MAKELPARAM(5, 5));
    SendMessageW(volumeUp, WM_LBUTTONUP, 0, MAKELPARAM(5, 5));
    check(app->volume == 50 + app->compactStep * 2,
          L"Rapid compact volume presses both count, including double-click second press");
    volume(0);
    command(main_cmdPause);
    check(app->status == 2, L"Pause button state");
    command(main_cmdPlay);
    check(app->status == 1, L"Play button state");
    command(main_cmdStop);
    check(app->status == 0 && app->realPosition == 0, L"Stop button state and position");
    command(main_cmdnext);
    check(selectedIndex() == 1 && app->status == 1, L"Next file");
    command(main_cmdnext);
    check(selectedIndex() == 0, L"Next wraps at folder end");
    command(main_cmdPrev);
    check(selectedIndex() == 1, L"Previous wraps at folder start");
    command(main_chkShuffle);
    command(main_cmdnext);
    check(app->checked[main_chkShuffle] && selectedIndex() >= 0 && selectedIndex() < 2,
          L"Shuffle stays inside current folder");
    command(main_chkShuffle);
    volume(140);
    check(app->volume == 100 && equal(readINI(app->presets, L"General", L"Volume"), L"100"),
          L"Volume upper limit and persistence");
    volume(-10);
    check(app->volume == 0, L"Volume lower limit");
    speed(150);
    check(app->speed == 150 && equal(windowText(control(main_lblRate)), L"x1.5"), L"Speed label and control");
    setText(main_txtStartTime, L"00:02");
    setText(main_txtEndTime, L"00:04");
    command(main_cmdApply);
    check(app->repeating && app->start == 2 && app->end == 4 && !IsWindowEnabled(control(main_txtStartTime)),
          L"Enable segment repeat and lock time fields");
    check(equal(readINI(app->repeatlist, lower(app->current), L"Start"), L"2") &&
              equal(readINI(app->repeatlist, lower(app->current), L"End"), L"4"),
          L"Persist segment by lowercase Unicode path");
    seek(4.2);
    tick();
    check(app->realPosition < 3 && app->realPosition >= 2, L"Segment boundary seeks to start");
    command(main_cmdApply);
    check(!app->repeating && IsWindowEnabled(control(main_txtStartTime)),
          L"Disable segment repeat and unlock time fields");
    command(main_chkRepeat);
    seek(app->playback.duration());
    app->ended = true;
    tick();
    check(app->realPosition == 0 && !app->ended, L"Whole-track repeat restarts same track");
    command(main_chkRepeat);
    int before = selectedIndex();
    app->ended = true;
    tick();
    check(selectedIndex() != before, L"Natural completion advances when repeat is off");
    command(menu_mnuShowTrackInCaption);
    check(equal(windowText(app->main), leaf(app->current)), L"Track in window caption");
    command(menu_mnuShowTrackInCaption);
    check(equal(windowText(app->main), title), L"Restore application caption");
    command(menu_mnuCloseToCompact);
    command(menu_mnuMinimizeToCompact);
    check(app->closeCompact && app->minimizeCompact && equal(readINI(app->presets, L"View", L"CloseToCompact"), L"1"),
          L"Compact preferences persist");
    loadFile(Text(L"http://127.0.0.1:52751/test.wav"));
    check((GetMenuState(app->menu, menu_mnuReverse, MF_BYCOMMAND) & MF_GRAYED) &&
              (GetMenuState(app->menu, menu_mnuRepeat, MF_BYCOMMAND) & MF_GRAYED),
          L"Stream disables reverse and repeat menu items with matching controls");
    command(menu_mnuReverse);
    command(menu_mnuRepeat);
    check(!app->reverseEnabled && !app->checked[main_chkRepeat], L"Disabled stream mode actions do not change state");
    check(equal(app->playback.component, L"Media Foundation") && !app->playback.supportsSpeed,
          L"Fallback backend has no vinyl speed support");
    double streamPosition = app->realPosition;
    check(arrowKey(control(main_Drive1), VK_RIGHT) && app->realPosition == streamPosition,
          L"Stream arrow key does not seek or fall through to drive navigation");
    check(app->speed == 100 && equal(windowText(control(main_lblRate)), L"x1"), L"Fallback backend resets speed to x1");
    check(!IsWindowEnabled(control(main_hsRate)), L"Fallback backend disables speed scrollbar");
    RECT rateRect{};
    GetClientRect(control(main_hsRate), &rateRect);
    HDC rateDC = CreateCompatibleDC(nullptr);
    BITMAPINFO bitmapInfo{};
    bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmapInfo.bmiHeader.biWidth = rateRect.right;
    bitmapInfo.bmiHeader.biHeight = -rateRect.bottom;
    bitmapInfo.bmiHeader.biPlanes = 1;
    bitmapInfo.bmiHeader.biBitCount = 32;
    bitmapInfo.bmiHeader.biCompression = BI_RGB;
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(rateDC, &bitmapInfo, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ previousBitmap = SelectObject(rateDC, bitmap);
    SendMessageW(control(main_hsRate), WM_PRINTCLIENT, reinterpret_cast<WPARAM>(rateDC), PRF_CLIENT);
    check(GetPixel(rateDC, rateRect.right / 2, rateRect.bottom / 2) == GetSysColor(COLOR_SCROLLBAR),
          L"Disabled speed paints inactive track without an active-looking thumb");
    COLORREF activeArrow = GetSysColor(COLOR_BTNTEXT);
    int darkArrowPixels = 0;
    int arrowWidth = GetSystemMetrics(SM_CXHSCROLL);
    for (int y = 4; y < rateRect.bottom - 4; ++y)
        for (int x = 4; x < arrowWidth - 4; ++x)
            if (GetPixel(rateDC, x, y) == activeArrow)
                ++darkArrowPixels;
    check(darkArrowPixels == 0, L"Disabled speed arrow uses inactive colors");
    SelectObject(rateDC, previousBitmap);
    DeleteObject(bitmap);
    DeleteDC(rateDC);
    SCROLLBARINFO rateInfo{};
    rateInfo.cbSize = sizeof(rateInfo);
    BOOL rateInfoAvailable = GetScrollBarInfo(control(main_hsRate), OBJID_CLIENT, &rateInfo);
    check(rateInfoAvailable && (rateInfo.rgstate[0] & STATE_SYSTEM_UNAVAILABLE) &&
              (rateInfo.rgstate[1] & STATE_SYSTEM_UNAVAILABLE) && (rateInfo.rgstate[5] & STATE_SYSTEM_UNAVAILABLE),
          L"Fallback disables entire speed scrollbar and both arrows");
    SendMessageW(app->main, WM_HSCROLL, MAKEWPARAM(SB_LINERIGHT, 0), (LPARAM)control(main_hsRate));
    check(app->speed == 100 && equal(windowText(control(main_lblRate)), L"x1"),
          L"Disabled speed scrollbar ignores scroll commands");
    selectTrack(0);
    check(app->playback.supportsSpeed && IsWindowEnabled(control(main_hsRate)),
          L"DirectShow restores speed scrollbar interaction");
    check(GetScrollBarInfo(control(main_hsRate), OBJID_CLIENT, &rateInfo) &&
              !(rateInfo.rgstate[1] & STATE_SYSTEM_UNAVAILABLE) && !(rateInfo.rgstate[3] & STATE_SYSTEM_UNAVAILABLE) &&
              !(rateInfo.rgstate[5] & STATE_SYSTEM_UNAVAILABLE),
          L"DirectShow restores speed arrows and thumb");
    SendMessageW(app->main, WM_HSCROLL, MAKEWPARAM(SB_LINERIGHT, 0), (LPARAM)control(main_hsRate));
    check(app->speed == 101, L"DirectShow speed accepts scroll commands after fallback");
    auto mediaPath = static_cast<wchar_t*>(allocate(65536));
    if (GetEnvironmentVariableW(L"STARDISK_MEDIA_TEST_FILE", mediaPath, 32768))
    {
        loadFile(Text(mediaPath));
        check(equal(app->playback.component, L"Media Foundation") && !app->isStream,
              L"Supplied local file uses Media Foundation");
        check(app->playback.duration() > 1 && SendMessageW(control(main_hsprogress), TBM_GETRANGEMAX, 0, 0) > 1,
              L"Local Media Foundation duration sets nonzero seek range");
        check(equal(windowText(control(main_lblTiming)), Text(L"00:00 / ") + timeText(int(app->playback.duration()))),
              L"Local Media Foundation duration appears in timing label");
        Text actualInfo = windowText(control(main_lblSysInfo));
        setText(main_lblSysInfo,
                L"abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789");
        check(paintedTextRows(control(main_lblSysInfo)) >= 3,
              L"Long filename without spaces actually paints on multiple title rows");
        setText(main_lblSysInfo, actualInfo);
        seek(10);
        check(app->realPosition > 9.9 && app->realPosition < 10.1,
              L"Known duration no longer clamps local Media Foundation seek to zero");
    }
    freeMemory(mediaPath);
    clearMedia(true);
    check(!app->playback.hasMedia() && app->status == 0 && app->current.empty(),
          L"Clear playback before file operation");
    for (int id :
         {main_cmdPlay, main_cmdPause, main_cmdStop, main_cmdPrev, main_cmdnext, main_chkRepeat, main_chkShuffle})
        check(equal(definition(id)->font, L"Webdings"), L"Original transport symbol font preserved");
    DestroyWindow(app->compact);
    DestroyWindow(app->main);
    for (auto font : app->fonts)
        if (font)
            DeleteObject(font);
    app->~App();
    freeMemory(app);
    CoUninitialize();
    log(Text(L"Failures: ") + number(failures));
    CloseHandle(testLog);
    ExitProcess(failures ? 1 : 0);
}
