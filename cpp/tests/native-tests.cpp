#include "../src/playback.h"

namespace
{
HANDLE logFile;
int failures = 0;
void log(const Text& text)
{
    Text line = text + L"\r\n";
    int count = WideCharToMultiByte(CP_UTF8, 0, line, line.size(), nullptr, 0, nullptr, nullptr);
    auto bytes = static_cast<char*>(allocate(count));
    WideCharToMultiByte(CP_UTF8, 0, line, line.size(), bytes, count, nullptr, nullptr);
    DWORD written;
    WriteFile(logFile, bytes, count, &written, nullptr);
    freeMemory(bytes);
}
void check(bool result, const wchar_t* label)
{
    log(Text(result ? L"PASS " : L"FAIL ") + label);
    if (!result)
        ++failures;
}
void checkHR(HRESULT result, const wchar_t* label)
{
    if (FAILED(result))
        log(hexCode(result) + L" " + windowsError(result));
    check(SUCCEEDED(result), label);
}
void delay(unsigned ms)
{
    ULONGLONG start = GetTickCount64();
    do
    {
        pumpMessages();
        Sleep(5);
    } while (GetTickCount64() - start < ms);
}
void checkRates(Playback& player)
{
    const int rates[] = {50, 100, 150, 200};
    for (int rate : rates)
    {
        log(Text(L"Testing speed: ") + number(rate) + L"%");
        checkHR(player.pause(), L"Pause before speed change");
        checkHR(player.rate(rate), L"Set playback speed");
        checkHR(player.seek(3.5), L"Seek at changed speed");
        check(player.position() > 3.4 && player.position() < 3.6, L"Speed does not scale seek target");
        checkHR(player.run(), L"Run at changed speed");
        delay(300);
        double before = player.position();
        ULONGLONG start = GetTickCount64();
        delay(2000);
        double advanced = player.position() - before;
        double expected = double(GetTickCount64() - start) * rate / 100000.0;
        log(Text(L"Advanced milliseconds: ") + number(int(advanced * 1000)) + L", expected: " +
            number(int(expected * 1000)));
        check(advanced > expected - 0.20 && advanced < expected + 0.20,
              L"Media position advances at selected speed, not speed squared");
        before = player.position();
        int changedRate = rate == 200 ? 50 : 200;
        checkHR(player.rate(changedRate), L"Change speed while playing");
        double afterChange = player.position();
        check(afterChange > before - 0.05 && afterChange < before + 0.15,
              L"Changing speed preserves precise media position");
        start = GetTickCount64();
        delay(400);
        advanced = player.position() - afterChange;
        expected = double(GetTickCount64() - start) * changedRate / 100000.0;
        check(advanced > expected - 0.15 && advanced < expected + 0.15,
              L"Running clock uses new speed immediately");
        checkHR(player.rate(rate), L"Restore tested speed while playing");
        checkHR(player.pause(), L"Pause at changed speed");
        delay(100);
        before = player.position();
        delay(200);
        double drift = player.position() - before;
        check(drift > -0.05 && drift < 0.05, L"Changed speed clock holds on pause");
        checkHR(player.seek(player.duration() - 6.0), L"Seek near end at changed speed");
        bool ended = false;
        player.pollEnded(ended); // Drain previous graph events.
        ended = false;
        start = GetTickCount64();
        checkHR(player.run(), L"Run last segment");
        while (!ended && GetTickCount64() - start < 14000)
        {
            delay(10);
            player.pollEnded(ended);
        }
        double elapsed = double(GetTickCount64() - start) / 1000.0;
        expected = 600.0 / rate;
        log(Text(L"End milliseconds: ") + number(int(elapsed * 1000)) + L", expected: " +
            number(int(expected * 1000)) + L", position: " + number(int(player.position() * 1000)) +
            L", duration: " + number(int(player.duration() * 1000)));
        check(ended && elapsed > expected - 0.35 && elapsed < expected + 0.45,
              L"Natural end agrees with selected speed and media position");
        check(player.position() > player.duration() - 0.2, L"Position reaches duration at natural end");
        checkHR(player.stop(), L"Stop after speed test");
    }
    checkHR(player.rate(100), L"Restore normal speed");
    checkHR(player.seek(0), L"Restore starting position");
}
} // namespace

extern "C" void testEntry()
{
    auto path = static_cast<wchar_t*>(allocate(65536));
    GetModuleFileNameW(nullptr, path, 32768);
    Text root = parent(parent(Text(path)));
    freeMemory(path);
    Text qa = join(root, L"qa"), ini = join(qa, L"native-tests.ini");
    logFile = CreateFileW(join(qa, L"native-tests.log"), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                          FILE_ATTRIBUTE_NORMAL, nullptr);
    if (logFile == INVALID_HANDLE_VALUE)
        ExitProcess(100);
    check(equal(parent(Text(L"C:\\")), L"C:\\"), L"Drive root remains intact");
    check(equal(parent(Text(L"C:\\Музыка\\ÄÖÜ")), L"C:\\Музыка"), L"Unicode parent directory");
    check(equal(join(Text(L"C:\\"), Text(L"тест.wav")), L"C:\\тест.wav"), L"Join drive root without duplicate slash");
    check(equal(trim(Text(L" \t ÄÖÜ тест \r\n")), L"ÄÖÜ тест"), L"Unicode whitespace trimming");
    int parsed = 0;
    check(integer(L"32767", parsed) && parsed == 32767 && !integer(L"2147483648", parsed), L"INI numeric bounds");
    check(equal(timeText(6001), L"100:01"), L"Long durations are not truncated");
    Text url = L"https://example.invalid/ÄÖÜ/поток?x=1&y=\\u0041\U0001F3B5";
    check(equal(decodeStream(encodeStream(url)), url),
          L"Stream URLs preserve Unicode, surrogate pairs, query strings and backslashes");
    check(equal(decodeStream(Text(L"\\u00ZZ")), L"\\u00ZZ"), L"Malformed legacy escapes are preserved");
    check(ensureINI(ini), L"Create Unicode INI");
    check(writeINI(ini, L"General", L"LastPath", L"D:\\Музыка\\ÄÖÜ"), L"Write Unicode path");
    check(equal(readINI(ini, L"General", L"LastPath"), L"D:\\Музыка\\ÄÖÜ"), L"Read Unicode path");
    check(writeINI(ini, L"Streams", L"URL1", encodeStream(url)) &&
              equal(decodeStream(readINI(ini, L"Streams", L"URL1")), url),
          L"Round-trip escaped stream URL through Windows profile API");
    TextList list;
    for (int i = 0; i < 100; ++i)
        list.add(number(i));
    check(list.count == 100 && equal(list[99], L"99") && list.contains(Text(L"51")), L"Dynamic stream list growth");
    HRESULT init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    checkHR(init, L"Initialize system COM");
    if (SUCCEEDED(init))
    {
        Playback player;
        Text attempts;
        Text wave = join(qa, L"native-audio\\01 - ÄÖÜ тест.wav");
        HRESULT hr = player.open(wave, attempts);
        checkHR(hr, L"Open Unicode WAV path using system playback");
        if (FAILED(hr))
            log(attempts);
        else
        {
            log(Text(L"Backend: ") + player.component);
            check(equal(player.component, L"DirectShow") && player.supportsSpeed,
                  L"DirectShow is preferred and enables speed");
            double d = player.duration();
            check(d > 29.9 && d < 30.1, L"Read exact media duration");
            checkHR(player.volume(0), L"Mute");
            checkHR(player.volume(100), L"Full volume");
            checkHR(player.volume(25), L"Logarithmic volume mapping");
            checkHR(player.seek(3.5), L"Fractional seek");
            check(player.position() > 3.4 && player.position() < 3.6, L"Seek position");
            checkHR(player.run(), L"Play");
            delay(250);
            check(player.position() > 3.6, L"Playback clock advances");
            checkHR(player.pause(), L"Pause");
            delay(50);
            double paused = player.position();
            delay(150);
            check(player.position() - paused < 0.05, L"Pause holds playback clock");
            checkHR(player.rate(150), L"Vinyl speed adjustment");
            checkHR(player.run(), L"Resume");
            delay(200);
            check(player.position() > paused + 0.20, L"Speed changes playback clock");
            checkHR(player.stop(), L"Stop");
            checkHR(player.seek(0), L"Stop resets position");
            check(player.position() < 0.05, L"Position after stop");
            checkHR(player.volume(0), L"Mute speed tests");
            checkRates(player);
            player.close();
            check(!player.hasMedia(), L"Release all playback references");
            // Rename while no graph owns the file, then put the fixture back.
            Text temporary = join(qa, L"native-audio\\01 - temporary.wav");
            BOOL renamed = MoveFileExW(wave, temporary, 0);
            check(!!renamed, L"Media handle released for rename");
            if (renamed)
                check(!!MoveFileExW(temporary, wave, 0), L"Restore renamed fixture");
            hr = player.open(wave, attempts);
            checkHR(hr, L"Reopen after file operation");
            if (SUCCEEDED(hr))
            {
                checkHR(player.seek(29.8), L"Seek near end");
                checkHR(player.run(), L"Play end of file");
                bool ended = false;
                ULONGLONG started = GetTickCount64();
                while (!ended && GetTickCount64() - started < 3000)
                {
                    delay(25);
                    player.pollEnded(ended);
                }
                check(ended, L"Detect natural end of track");
            }
            player.close();
        }
        hr = player.open(L"http://127.0.0.1:52751/test.wav", attempts);
        checkHR(hr, L"Open HTTP stream via Media Foundation");
        if (FAILED(hr))
            log(attempts);
        else
        {
            check(equal(player.component, L"Media Foundation") && !player.supportsSpeed,
                  L"Stream backend disables vinyl speed");
            check(player.duration() > 29.9 && player.duration() < 30.1,
                  L"Media Foundation reads unsigned duration for finite media");
            checkHR(player.volume(25), L"Stream volume");
            checkHR(player.run(), L"Start HTTP stream");
            delay(150);
            checkHR(player.pause(), L"Pause HTTP stream");
            checkHR(player.run(), L"Resume HTTP stream");
            checkHR(player.stop(), L"Stop HTTP stream");
        }
        player.close();
        auto rateFile = static_cast<wchar_t*>(allocate(65536));
        if (GetEnvironmentVariableW(L"STARDISK_RATE_TEST_FILE", rateFile, 32768))
        {
            hr = player.open(Text(rateFile), attempts);
            checkHR(hr, L"Open additional compressed audio speed fixture");
            if (SUCCEEDED(hr))
            {
                check(equal(player.component, L"DirectShow"), L"Compressed fixture uses DirectShow");
                checkHR(player.volume(0), L"Mute compressed fixture");
                checkRates(player);
            }
            else
                log(attempts);
            player.close();
        }
        freeMemory(rateFile);
        check(equal(trackDisplay(wave), leaf(wave)), L"Missing tags fall back to Unicode filename");
        CoUninitialize();
    }
    log(Text(L"Failures: ") + number(failures));
    CloseHandle(logFile);
    ExitProcess(failures ? 1 : 0);
}
