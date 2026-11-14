#include "playback.h"
#include <propsys.h>
#include <propkey.h>

struct Playback::Events final : IMFPMediaPlayerCallback
{
    volatile LONG refs = 1, completed = 0, error = S_OK, ended = 0;
    HANDLE signal = CreateEventW(nullptr, FALSE, FALSE, nullptr);
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (IsEqualIID(iid, IID_IUnknown) || IsEqualIID(iid, __uuidof(IMFPMediaPlayerCallback)))
        {
            *out = this;
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&refs);
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        LONG n = InterlockedDecrement(&refs);
        if (!n)
        {
            CloseHandle(signal);
            this->~Events();
            freeMemory(this);
        }
        return n;
    }
    void STDMETHODCALLTYPE OnMediaPlayerEvent(MFP_EVENT_HEADER* e) override
    {
        if (FAILED(e->hrEvent))
            InterlockedExchange(&error, e->hrEvent);
        if (e->eEventType >= 0 && e->eEventType <= 12)
            InterlockedOr(&completed, 1L << e->eEventType);
        if (e->eEventType == MFP_EVENT_TYPE_PLAYBACK_ENDED)
            InterlockedExchange(&ended, 1);
        SetEvent(signal);
    }
};

void Playback::close()
{
    if (reverse_) reverseApi_.close(reverse_);
    reverse_ = nullptr;
    reverseEnded_ = false;
    if (control_)
        control_->Stop();
    if (mf_)
        mf_->Shutdown();
    release(graphEvents_);
    release(audio_);
    release(position_);
    release(control_);
    release(graph_);
    release(mf_);
    release(events_);
    // The MF module stays loaded until exit; callbacks may still be unwinding.
    component = L"";
    supportsSpeed = false;
    graphPosition_ = 0;
    graphRate_ = 1;
    graphStarted_ = 0;
    graphRunning_ = false;
}

HRESULT Playback::openGraph(const Text& path)
{
    HRESULT hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&graph_));
    if (FAILED(hr))
        return hr;
    hr = graph_->RenderFile(path, nullptr);
    if (SUCCEEDED(hr))
        hr = graph_->QueryInterface(IID_PPV_ARGS(&control_));
    if (SUCCEEDED(hr))
        hr = graph_->QueryInterface(IID_IMediaPosition, reinterpret_cast<void**>(&position_));
    if (SUCCEEDED(hr))
        hr = graph_->QueryInterface(IID_IBasicAudio, reinterpret_cast<void**>(&audio_));
    if (SUCCEEDED(hr))
        hr = graph_->QueryInterface(IID_IMediaEvent, reinterpret_cast<void**>(&graphEvents_));
    if (SUCCEEDED(hr))
    {
        component = L"DirectShow";
        supportsSpeed = true;
    }
    return hr;
}

HRESULT Playback::beginEvent(int type)
{
    InterlockedAnd(&events_->completed, ~(1L << type));
    return InterlockedCompareExchange(&events_->error, 0, 0);
}
HRESULT Playback::waitEvent(int type)
{
    ULONGLONG start = GetTickCount64();
    while (!(InterlockedCompareExchange(&events_->completed, 0, 0) & (1L << type)))
    {
        HRESULT hr = InterlockedCompareExchange(&events_->error, 0, 0);
        if (FAILED(hr))
            return hr;
        if (GetTickCount64() - start > 10000)
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        MsgWaitForMultipleObjects(1, &events_->signal, FALSE, 20, QS_ALLINPUT);
        pumpMessages();
    }
    return InterlockedCompareExchange(&events_->error, 0, 0);
}
HRESULT Playback::open(const Text& path, Text& attempts, bool reverse)
{
    close();
    busy = true;
    HRESULT hr = E_FAIL;
    attempts = L"";
    if (reverse)
    {
        if (streamURL(path) || !reverseApi_.load()) { busy = false; return E_NOTIMPL; }
        hr = reverseApi_.open(path, 2, 0.25, &reverse_);
        if (SUCCEEDED(hr)) component = L"Media Foundation (реверс)";
        else attempts = reverseError();
        busy = false;
        return hr;
    }
    if (!streamURL(path))
    {
        hr = openGraph(path);
        if (SUCCEEDED(hr))
        {
            busy = false;
            return hr;
        }
        attempts = Text(L"Обычный DirectShow: ") + hexCode(hr) + L" — " + windowsError(hr);
        close();

    }
    if (!mfModule_)
        mfModule_ = LoadLibraryW(L"mfplay.dll");
    using CreatePlayer =
        HRESULT(WINAPI*)(LPCWSTR, BOOL, MFP_CREATION_OPTIONS, IMFPMediaPlayerCallback*, HWND, IMFPMediaPlayer**);
    auto create =
        mfModule_ ? reinterpret_cast<CreatePlayer>(GetProcAddress(mfModule_, "MFPCreateMediaPlayer")) : nullptr;
    if (!create)
        hr = HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    else
    {
        events_ = new (allocate(sizeof(Events))) Events();
        hr = beginEvent(MFP_EVENT_TYPE_MEDIAITEM_SET);
        bool stream = streamURL(path);
        if (SUCCEEDED(hr))
            hr = create(stream ? path.c_str() : nullptr, FALSE, MFP_OPTION_FREE_THREADED_CALLBACK, events_, nullptr,
                        &mf_);
        IMFPMediaItem* item = nullptr;
        if (SUCCEEDED(hr) && !stream)
            hr = mf_->CreateMediaItemFromURL(path, TRUE, 0, &item);
        if (SUCCEEDED(hr) && !stream)
            hr = mf_->SetMediaItem(item);
        release(item);
        if (SUCCEEDED(hr))
            hr = waitEvent(MFP_EVENT_TYPE_MEDIAITEM_SET);
    }
    if (FAILED(hr))
    {
        if (!attempts.empty())
            attempts += L"\r\n";
        attempts += Text(L"Media Foundation: ") + hexCode(hr) + L" — " + windowsError(hr);
        close();
    }
    else
    {
        component = L"Media Foundation";
        supportsSpeed = false;
    }
    busy = false;
    return hr;
}
HRESULT Playback::run()
{
    if (reverse_) { reverseEnded_ = false; return reverseApi_.play(reverse_); }
    if (control_)
    {
        HRESULT hr = control_->Run();
        if (SUCCEEDED(hr) && !graphRunning_)
        {
            graphStarted_ = GetTickCount64();
            graphRunning_ = true;
        }
        return hr;
    }
    if (!mf_)
        return S_FALSE;
    MFP_MEDIAPLAYER_STATE state{};
    HRESULT hr = mf_->GetState(&state);
    if (FAILED(hr) || state == MFP_MEDIAPLAYER_STATE_PLAYING)
        return hr;
    InterlockedExchange(&events_->ended, 0);
    hr = beginEvent(MFP_EVENT_TYPE_PLAY);
    if (SUCCEEDED(hr))
        hr = mf_->Play();
    if (SUCCEEDED(hr))
    {
        busy = true;
        hr = waitEvent(MFP_EVENT_TYPE_PLAY);
        busy = false;
    }
    return hr;
}
HRESULT Playback::pause()
{
    if (reverse_) return reverseApi_.pause(reverse_);
    if (control_)
    {
        double current = position();
        HRESULT hr = control_->Pause();
        if (SUCCEEDED(hr))
        {
            graphPosition_ = current;
            graphRunning_ = false;
        }
        return hr;
    }
    if (!mf_)
        return S_FALSE;
    MFP_MEDIAPLAYER_STATE state{};
    HRESULT hr = mf_->GetState(&state);
    if (FAILED(hr) || state != MFP_MEDIAPLAYER_STATE_PLAYING)
        return hr;
    hr = beginEvent(MFP_EVENT_TYPE_PAUSE);
    if (SUCCEEDED(hr))
        hr = mf_->Pause();
    if (SUCCEEDED(hr))
    {
        busy = true;
        hr = waitEvent(MFP_EVENT_TYPE_PAUSE);
        busy = false;
    }
    return hr;
}
HRESULT Playback::stop()
{
    if (reverse_) { reverseEnded_ = false; return reverseApi_.stop(reverse_); }
    if (control_)
    {
        double current = position();
        HRESULT hr = control_->Stop();
        if (SUCCEEDED(hr))
        {
            graphPosition_ = current;
            graphRunning_ = false;
        }
        return hr;
    }
    if (!mf_)
        return S_FALSE;
    MFP_MEDIAPLAYER_STATE state{};
    HRESULT hr = mf_->GetState(&state);
    if (FAILED(hr) || state <= MFP_MEDIAPLAYER_STATE_STOPPED)
        return hr;
    hr = beginEvent(MFP_EVENT_TYPE_STOP);
    if (SUCCEEDED(hr))
        hr = mf_->Stop();
    if (SUCCEEDED(hr))
    {
        busy = true;
        hr = waitEvent(MFP_EVENT_TYPE_STOP);
        busy = false;
    }
    InterlockedExchange(&events_->ended, 0);
    return hr;
}
double Playback::duration() const
{
    if (reverse_) return reverseApi_.duration(reverse_);
    double seconds = 0;
    if (position_)
        position_->get_Duration(&seconds);
    else if (mf_)
    {
        PROPVARIANT value{};
        if (SUCCEEDED(mf_->GetDuration(MFP_POSITIONTYPE_100NS, &value)))
        {
            // IMFPMediaPlayer returns an unsigned duration (VT_UI8), unlike
            // IMFPMediaItem::GetDuration. Accept the signed form as well.
            if (value.vt == VT_UI8)
                seconds = double(value.uhVal.QuadPart) / 10000000.0;
            else if (value.vt == VT_I8)
                seconds = double(value.hVal.QuadPart) / 10000000.0;
        }
        PropVariantClear(&value);
    }
    return seconds;
}
double Playback::position() const
{
    if (reverse_) return reverseApi_.position(reverse_);
    double seconds = 0;
    if (position_)
    {
        seconds = graphPosition_;
        if (graphRunning_)
            seconds += double(GetTickCount64() - graphStarted_) * graphRate_ / 1000.0;
        double length = duration();
        if (length > 0 && seconds > length)
            seconds = length;
    }
    else if (mf_)
    {
        PROPVARIANT value{};
        if (SUCCEEDED(mf_->GetPosition(MFP_POSITIONTYPE_100NS, &value)) && value.vt == VT_I8)
            seconds = double(value.hVal.QuadPart) / 10000000.0;
        PropVariantClear(&value);
    }
    return seconds;
}
HRESULT Playback::seek(double seconds)
{
    if (reverse_) { reverseEnded_ = false; return reverseApi_.seek(reverse_, seconds); }
    if (position_)
    {
        HRESULT hr = position_->put_CurrentPosition(seconds);
        if (SUCCEEDED(hr))
        {
            // Running graph position is unreliable at non-unit speed.
            graphPosition_ = seconds;
            graphStarted_ = GetTickCount64();
        }
        return hr;
    }
    if (!mf_)
        return S_FALSE;
    PROPVARIANT value{};
    value.vt = VT_I8;
    value.hVal.QuadPart = LONGLONG(seconds * 10000000.0);
    InterlockedExchange(&events_->ended, 0);
    HRESULT hr = beginEvent(MFP_EVENT_TYPE_POSITION_SET);
    if (SUCCEEDED(hr))
        hr = mf_->SetPosition(MFP_POSITIONTYPE_100NS, &value);
    if (SUCCEEDED(hr))
    {
        busy = true;
        hr = waitEvent(MFP_EVENT_TYPE_POSITION_SET);
        busy = false;
    }
    return hr;
}
HRESULT Playback::volume(int percent)
{
    percent = clamp(percent, 0, 100);
    if (reverse_) return reverseApi_.volume(reverse_, percent);
    if (mf_)
        return mf_->SetVolume(float(percent) / 100.0f);
    if (!audio_)
        return S_FALSE;
    // 2000*log10(percent/100), as in the original; no math runtime.
    LONG db = -10000;
    if (percent)
    {
        double x = double(percent) / 100.0;
        int exponent = 0;
        while (x < 1.0)
        {
            x *= 2.0;
            --exponent;
        }
        double z = (x - 1.0) / (x + 1.0), term = z, sum = 0;
        for (int k = 1; k < 40; k += 2)
        {
            sum += term / k;
            term *= z * z;
        }
        db = LONG(2000.0 * (2.0 * sum + exponent * 0.6931471805599453) / 2.302585092994046 - 0.5);
    }
    return audio_->put_Volume(db);
}
HRESULT Playback::rate(int percent)
{
    if (!supportsSpeed || !position_)
        return S_FALSE;
    double current = position();
    double rate = double(percent) / 100.0;
    HRESULT hr = position_->put_Rate(rate);
    if (SUCCEEDED(hr))
    {
        graphRate_ = rate;
        // SetRate can re-seek using the graph's incorrectly scaled position.
        // Restore the actual media position using our clock instead.
        hr = seek(current);
    }
    return hr;
}
HRESULT Playback::pollEnded(bool& ended)
{
    ended = false;
    if (reverse_)
    {
        int state = reverseApi_.state(reverse_);
        ended = state == 3 && !reverseEnded_;
        if (ended) reverseEnded_ = true;
        return state == 4 ? E_FAIL : S_OK;
    }
    if (graphEvents_)
    {
        LONG code;
        LONG_PTR a, b;
        while (graphEvents_->GetEvent(&code, &a, &b, 0) == S_OK)
        {
            if (code == EC_COMPLETE)
            {
                ended = true;
                graphPosition_ = duration();
                graphRunning_ = false;
            }
            graphEvents_->FreeEventParams(code, a, b);
        }
    }
    else if (events_)
    {
        ended = !!InterlockedExchange(&events_->ended, 0);
        return InterlockedCompareExchange(&events_->error, 0, 0);
    }
    return S_OK;
}

Text Playback::reverseError() const
{
    wchar_t message[512]{};
    if (reverseApi_.error) reverseApi_.error(reverse_, message, 512);
    return Text(message);
}
void Playback::swap(Playback& other)
{
    auto exchange = [](auto& a, auto& b) { auto value = a; a = b; b = value; };
    exchange(graph_, other.graph_); exchange(control_, other.control_);
    exchange(position_, other.position_); exchange(audio_, other.audio_);
    exchange(graphEvents_, other.graphEvents_); exchange(mf_, other.mf_);
    exchange(events_, other.events_); exchange(mfModule_, other.mfModule_);
    exchange(graphPosition_, other.graphPosition_); exchange(graphRate_, other.graphRate_);
    exchange(graphStarted_, other.graphStarted_); exchange(graphRunning_, other.graphRunning_);
    exchange(reverseApi_, other.reverseApi_); exchange(reverse_, other.reverse_);
    exchange(reverseEnded_, other.reverseEnded_); exchange(component, other.component);
    exchange(supportsSpeed, other.supportsSpeed); exchange(busy, other.busy);
}

static Text shellText(IPropertyStore* store, REFPROPERTYKEY key)
{
    PROPVARIANT value{};
    Text result;
    if (SUCCEEDED(store->GetValue(key, &value)))
    {
        if (value.vt == (VT_VECTOR | VT_LPWSTR))
        {
            for (ULONG i = 0; i < value.calpwstr.cElems; ++i)
            {
                Text part = trim(Text(value.calpwstr.pElems[i]));
                if (!part.empty())
                {
                    if (!result.empty())
                        result += L"; ";
                    result += part;
                }
            }
        }
        else
        {
            PWSTR p = nullptr;
            if (SUCCEEDED(PropVariantToStringAlloc(value, &p)))
            {
                result = p;
                CoTaskMemFree(p);
            }
        }
    }
    PropVariantClear(&value);
    wchar_t* p = const_cast<wchar_t*>(result.c_str());
    for (int i = 0; i < result.size(); ++i)
        if (p[i] == L'\r' || p[i] == L'\n' || p[i] == L'\t')
            p[i] = L' ';
    return trim(result);
}
Text trackDisplay(const Text& file)
{
    IPropertyStore* store = nullptr;
    Text result = leaf(file);
    if (SUCCEEDED(SHGetPropertyStoreFromParsingName(file, nullptr, GPS_HANDLERPROPERTIESONLY, IID_PPV_ARGS(&store))))
    {
        Text artist = shellText(store, PKEY_Music_Artist), title = shellText(store, PKEY_Title);
        if (!artist.empty() && !title.empty())
            result = artist + L" - " + title;
        store->Release();
    }
    return result;
}
