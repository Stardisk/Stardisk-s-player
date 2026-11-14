#pragma once
#include "base.h"
#include <dshow.h>
#include <mfplay.h>
#include "reverse_api.h"
extern "C" HRESULT WINAPI PropVariantToStringAlloc(REFPROPVARIANT value, PWSTR* text);

template <class T> inline void release(T*& p)
{
    if (p)
    {
        p->Release();
        p = nullptr;
    }
}

class Playback
{
    struct Events;
    IGraphBuilder* graph_ = nullptr;
    IMediaControl* control_ = nullptr;
    IMediaPosition* position_ = nullptr;
    IBasicAudio* audio_ = nullptr;
    IMediaEvent* graphEvents_ = nullptr;
    IMFPMediaPlayer* mf_ = nullptr;
    Events* events_ = nullptr;
    HMODULE mfModule_ = nullptr;
    // Audio renderer clocks can already include the rate (notably for MP3).
    // Keep UI media time on a monotonic clock, independent of that graph clock.
    double graphPosition_ = 0, graphRate_ = 1;
    ULONGLONG graphStarted_ = 0;
    bool graphRunning_ = false;
    ReverseApi reverseApi_;
    void* reverse_ = nullptr;
    bool reverseEnded_ = false;
    HRESULT openGraph(const Text& path);
    HRESULT waitEvent(int type);
    HRESULT beginEvent(int type);

  public:
    ~Playback() { close(); reverseApi_.unload(); }
    bool reverseAvailable() { return reverseApi_.load(); }
    bool isReverse() const { return reverse_ != nullptr; }
    Text reverseError() const;
    void swap(Playback& other);
    const wchar_t* component = L"";
    bool supportsSpeed = false;
    bool busy = false;
    bool hasMedia() const
    {
        return control_ || mf_ || reverse_;
    }
    HRESULT open(const Text& path, Text& attempts, bool reverse = false);
    void close();
    HRESULT run();
    HRESULT pause();
    HRESULT stop();
    double duration() const;
    double position() const;
    HRESULT seek(double seconds);
    HRESULT volume(int percent);
    HRESULT rate(int percent);
    HRESULT pollEnded(bool& ended);
};
Text trackDisplay(const Text& file);
