#pragma once
// Shared decoding engine, moved from the verified PoC. All use lives in reverse.dll.
#define NOMINMAX
#include <windows.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <mmsystem.h>

#include <shlwapi.h>
#include <algorithm>
#include <array>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <stdexcept>
#include <string>
#include <vector>

// Temporary workaround for the observed Source Reader read-ahead/seek race.
// Build with REVERSE_CONTINUE_MEMORY_PROBE=0 to restore the original path.
#ifndef REVERSE_CONTINUE_MEMORY_PROBE
#define REVERSE_CONTINUE_MEMORY_PROBE 1
#endif
extern "C" HRESULT WINAPI MFCreateMFByteStreamOnStream(IStream*,IMFByteStream**);

using Bytes = std::vector<unsigned char>;
struct AudioFormat {
    unsigned hz=0, channels=0;
    unsigned align() const { return channels*2; }
    bool operator==(const AudioFormat& o) const { return hz==o.hz && channels==o.channels; }
};
void check(HRESULT hr, const char* what) {
    if (FAILED(hr)) { char s[200]; sprintf(s, "%s: HRESULT 0x%08lX", what, (unsigned long)hr); throw std::runtime_error(s); }
}
template<class T> void drop(T*& p) { if (p) p->Release(); p=nullptr; }
struct Reader {
    IMFSourceReader* r=nullptr;
    Bytes probePCM;
    bool continueProbe=false,probeEnded=false;
    long long total=0;
    long long initialGap=0;
    double mp3Quantum=0;
    unsigned nativeRate=0, nativeChannels=0;
    AudioFormat fmt;
    ~Reader() { drop(r); }
    void open(const wchar_t* path) {
        check(MFCreateSourceReaderFromURL(path,nullptr,&r),"Open input");
        configure();
    }
    void openMemory(const BYTE* data,UINT size) {
        IStream* stream=SHCreateMemStream(data,size);
        if(!stream) throw std::runtime_error("Create memory stream failed");
        IMFByteStream* bytes=nullptr;
        HRESULT hr=MFCreateMFByteStreamOnStream(stream,&bytes); drop(stream);
        check(hr,"MF memory stream");
        hr=MFCreateSourceReaderFromByteStream(bytes,nullptr,&r); drop(bytes);
        check(hr,"Open MP3 memory segment"); configure(REVERSE_CONTINUE_MEMORY_PROBE!=0);
    }
    void configure(bool keepProbe=false) {
        continueProbe=keepProbe;
        check(r->SetStreamSelection(MF_SOURCE_READER_ALL_STREAMS,FALSE),"Deselect");
        check(r->SetStreamSelection(MF_SOURCE_READER_FIRST_AUDIO_STREAM,TRUE),"Select audio");
        IMFMediaType* native=nullptr;
        if(SUCCEEDED(r->GetNativeMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM,0,&native))) {
            GUID subtype{}; UINT32 rate=0;
            native->GetGUID(MF_MT_SUBTYPE,&subtype); native->GetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND,&rate);
            nativeRate=rate; native->GetUINT32(MF_MT_AUDIO_NUM_CHANNELS,&nativeChannels);
            if(IsEqualGUID(subtype,MFAudioFormat_MP3) && rate) mp3Quantum=576;
            drop(native);
        }
        if(!nativeRate || !nativeChannels || nativeChannels>8)
            throw std::runtime_error("Missing or unsupported native audio format");
        fmt={nativeRate,nativeChannels};
        IMFMediaType* t=nullptr;
        check(MFCreateMediaType(&t),"Media type");
        t->SetGUID(MF_MT_MAJOR_TYPE,MFMediaType_Audio);
        t->SetGUID(MF_MT_SUBTYPE,MFAudioFormat_PCM);
        t->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS,fmt.channels);
        t->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND,fmt.hz);
        t->SetUINT32(MF_MT_AUDIO_BITS_PER_SAMPLE,16);
        t->SetUINT32(MF_MT_AUDIO_BLOCK_ALIGNMENT,fmt.align());
        t->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND,fmt.hz*fmt.align());
        HRESULT hr=r->SetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM,nullptr,t); drop(t);
        check(hr,"Select PCM16 at original sample rate and channel count");
        PROPVARIANT v; PropVariantInit(&v);
        check(r->GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE,MF_PD_DURATION,&v),"Duration");
        total=llround((double)v.uhVal.QuadPart*fmt.hz/10000000.0); PropVariantClear(&v);
        if(mp3Quantum) total=llround(round(total/mp3Quantum)*mp3Quantum);
        // Some Windows MP3 decoders trim priming samples from the first buffer
        // without moving timestamps of subsequent buffers. Map to contiguous PCM.
        Bytes a,b; long long ta=0,tb=0;
        if(next(a,ta) && next(b,tb) && !a.empty() && !b.empty()) {
            initialGap=std::max(0LL,tb-ta-(long long)a.size()/fmt.align());
            total-=initialGap;
        }
        if(continueProbe) {
            // Keep the two PCM probes instead of seeking this short stream back
            // while MF may already be requesting its next compressed sample.
            probePCM=std::move(a);
            probePCM.insert(probePCM.end(),b.begin(),b.end());
            return;
        }
        PROPVARIANT zero; PropVariantInit(&zero); zero.vt=VT_I8; zero.hVal.QuadPart=0;
        check(r->SetCurrentPosition(GUID_NULL,zero),"Rewind decoder");
    }
    bool next(Bytes& b,long long& first) {
        DWORD flags=0; LONGLONG time=0; IMFSample* s=nullptr;
        check(r->ReadSample(MF_SOURCE_READER_FIRST_AUDIO_STREAM,0,nullptr,&flags,&time,&s),"Decode");
        if(continueProbe) probeEnded=!!(flags & MF_SOURCE_READERF_ENDOFSTREAM);
        if(flags & MF_SOURCE_READERF_ERROR) { drop(s); throw std::runtime_error("Decoder stream error"); }
        if(flags & MF_SOURCE_READERF_CURRENTMEDIATYPECHANGED) { drop(s); throw std::runtime_error("Audio format changed"); }
        b.clear(); first=llround((double)time*fmt.hz/10000000.0);
        // MP3 source timestamps can drift by a few samples due to rounding.
        // MPEG Layer III output starts on 576-sample granule boundaries.
        if(mp3Quantum) first=llround(round(first/mp3Quantum)*mp3Quantum);
        if(s) {
            IMFMediaBuffer* buf=nullptr; HRESULT hr=s->ConvertToContiguousBuffer(&buf); drop(s); check(hr,"Sample buffer");
            BYTE* data=nullptr; DWORD size=0; hr=buf->Lock(&data,nullptr,&size);
            if(FAILED(hr)) { drop(buf); check(hr,"Lock buffer"); }
            b.assign(data,data+size); buf->Unlock(); drop(buf);
        }
        return !(flags & MF_SOURCE_READERF_ENDOFSTREAM) || !b.empty();
    }
    Bytes all() {
        Bytes out=std::move(probePCM),b; long long time;
        while(!(continueProbe && probeEnded) && next(b,time)) out.insert(out.end(),b.begin(),b.end());
        total=(long long)out.size()/fmt.align(); return out;
    }
    Bytes range(long long lo,long long hi) {
        // Decode forward with preroll, then trim by timestamps; never feed reversed MP3 to a decoder.
        PROPVARIANT v; PropVariantInit(&v); v.vt=VT_I8;
        v.hVal.QuadPart=std::max(0LL,lo+initialGap-fmt.hz)*10000000/fmt.hz;
        check(r->SetCurrentPosition(GUID_NULL,v),"Seek decoder");
        continueProbe=false;probeEnded=false;probePCM.clear();
        Bytes out((size_t)(hi-lo)*fmt.align()),b; long long time,filled=lo;
        while(next(b,time)) {
            if(time>0) time-=initialGap;
            long long end=time+(long long)b.size()/fmt.align();
            long long a=std::max(lo,time),z=std::min(hi,end);
            if(z>a) {
                if(a>filled) throw std::runtime_error("Gap in decoded timestamps; cannot promise seamless reverse");
                memcpy(out.data()+(a-lo)*fmt.align(),b.data()+(a-time)*fmt.align(),(size_t)(z-a)*fmt.align());
                filled=std::max(filled,z);
            }
            if(filled>=hi) break;
            if(time>=hi) break;
        }
        if(filled<hi) throw std::runtime_error("Decoder did not supply the requested range (duration/seek mismatch)");
        return out;
    }
};
void reversePCM(Bytes& b,unsigned Align) {
    size_t n=b.size()/Align;
    for(size_t i=0;i<n/2;i++) for(unsigned k=0;k<Align;k++) std::swap(b[i*Align+k],b[(n-1-i)*Align+k]);
}
Bytes readFile(const wchar_t* path) {
    FILE* f=_wfopen(path,L"rb"); if(!f) throw std::runtime_error("Cannot open file");
    _fseeki64(f,0,SEEK_END); auto n=_ftelli64(f); rewind(f);
    if(n<0 || n>1024LL*1024*1024) { fclose(f); throw std::runtime_error("Indexed MP3 input limit: 1 GiB"); }
    Bytes b((size_t)n); size_t got=fread(b.data(),1,b.size(),f); fclose(f);
    if(got!=b.size()) throw std::runtime_error("File read failed"); return b;
}
int frameSize(const unsigned char* p) {
    unsigned h=(unsigned(p[0])<<24)|(unsigned(p[1])<<16)|(unsigned(p[2])<<8)|p[3];
    if((h&0xffe00000)!=0xffe00000) return 0;
    int ver=(h>>19)&3,layer=(h>>17)&3,br=(h>>12)&15,sr=(h>>10)&3;
    if(ver==1 || layer!=1 || !br || br==15 || sr==3) return 0;
    static int rates[3]={44100,48000,32000};
    static int br1[16]={0,32,40,48,56,64,80,96,112,128,160,192,224,256,320,0};
    static int br2[16]={0,8,16,24,32,40,48,56,64,80,96,112,128,144,160,0};
    int hz=rates[sr]/(ver==3?1:ver==2?2:4);
    return (ver==3?144000:72000)*(ver==3?br1[br]:br2[br])/hz+((h>>9)&1);
}
struct IndexedMP3 {
    Bytes compressed;
    std::vector<std::pair<size_t,size_t>> frames;
    long long total=0;
    int samples=1152;
    AudioFormat fmt;
    void open(const wchar_t* path,long long delay,AudioFormat format) {
        fmt=format;
        compressed=readFile(path); auto& b=compressed; size_t p=0;
        if(b.size()>=10 && !memcmp(b.data(),"ID3",3)) {
            p=10+((b[6]&127)<<21)+((b[7]&127)<<14)+((b[8]&127)<<7)+(b[9]&127);
            if(b[5]&16) p+=10;
        }
        while(p+4<=b.size()) {
            int n=frameSize(b.data()+p);
            if(n>0 && p+n<=b.size()) {
                if(frames.empty()) {
                    samples=((b[p+1]>>3)&3)==3?1152:576;
                    bool metadata=false;
                    for(size_t j=p+4;j+4<=p+n;j++)
                        if(!memcmp(b.data()+j,"Xing",4)||!memcmp(b.data()+j,"VBRI",4)) metadata=true;
                    if(metadata) { p+=n; continue; }
                }
                frames.emplace_back(p,n); p+=n;
            } else { if(!frames.empty()) break; ++p; }
        }
        if(frames.empty()) throw std::runtime_error("Could not index MP3 frames");
        total=(long long)frames.size()*samples-delay;
        // MPEG-2/2.5 decoders can withhold a final frame. Determine the true
        // endpoint from a short tail decode instead of assuming all frames emit PCM.
        size_t tailCount=(fmt.hz+samples-1)/samples+4;
        size_t tailFirst=frames.size()>tailCount?frames.size()-tailCount:0;
        size_t tailOffset=frames[tailFirst].first;
        // Preserve trailing tags too: the system MP3 source's EOF handling can
        // differ between a bare final frame and the original tagged stream.
        size_t tailEnd=compressed.size();
        Reader tail; tail.openMemory(compressed.data()+tailOffset,(UINT)(tailEnd-tailOffset));
        if(!(tail.fmt==fmt)) throw std::runtime_error("MP3 tail audio format changed");
        Bytes tailPCM=tail.all();
        total=(long long)tailFirst*samples+(long long)tailPCM.size()/fmt.align();

    }
    Bytes range(long long lo,long long hi) {
        // Exact frame positions replace the VBR source's approximate time seeks.
        // Decode one second of preceding frames to restore reservoir/filter state.
        size_t first=(size_t)(std::max(0LL,lo-fmt.hz)/samples);
        size_t last=std::min(frames.size(),(size_t)((hi+samples-1)/samples)+2);
        size_t offset=frames[first].first,end=frames[last-1].first+frames[last-1].second;
        Reader segment; segment.openMemory(compressed.data()+offset,(UINT)(end-offset));
        if(!(segment.fmt==fmt)) throw std::runtime_error("MP3 segment audio format changed");
        Bytes decoded=segment.all();
        size_t a=(size_t)(lo-(long long)first*samples)*fmt.align(),z=(size_t)(hi-(long long)first*samples)*fmt.align();
        if(z>decoded.size()) throw std::runtime_error("Indexed segment too short");
        return Bytes(decoded.begin()+a,decoded.begin()+z);
    }
};
struct Wav {
    FILE* f=nullptr; unsigned bytes=0;
    AudioFormat fmt;
    void open(const wchar_t* p,AudioFormat format) {
        fmt=format;
        f=_wfopen(p,L"wb"); if(!f) throw std::runtime_error("Cannot create WAV");
        unsigned char h[44]={}; fwrite(h,1,44,f);
    }
    void write(const Bytes& b) {
        if(!f) return;
        if(b.size()>0xffffffffULL-36-bytes) throw std::runtime_error("WAV exceeds RIFF size limit");
        if(fwrite(b.data(),1,b.size(),f)!=b.size()) throw std::runtime_error("WAV write failed"); bytes+=(unsigned)b.size();
    }
    ~Wav() {
        if(!f) return;
        unsigned h[11]={0x46464952,bytes+36,0x45564157,0x20746d66,16,(fmt.channels<<16)|1,fmt.hz,fmt.hz*fmt.align(),(16u<<16)|fmt.align(),0x61746164,bytes};
        rewind(f); fwrite(h,1,44,f); fclose(f);
    }
};
struct Output {
    HWAVEOUT wave=nullptr;
    struct Slot { WAVEHDR h{}; Bytes data; bool prepared=false; };
    std::array<Slot,3> slots;
    void open(AudioFormat format) {
        WAVEFORMATEXTENSIBLE ex{};
        ex.Format={WAVE_FORMAT_PCM,(WORD)format.channels,format.hz,format.hz*format.align(),(WORD)format.align(),16,0};
        if(format.channels>2) {
            ex.Format.wFormatTag=WAVE_FORMAT_EXTENSIBLE;
            ex.Format.cbSize=22; ex.Samples.wValidBitsPerSample=16; ex.SubFormat=MFAudioFormat_PCM;
        }
        if(waveOutOpen(&wave,WAVE_MAPPER,&ex.Format,0,0,CALLBACK_NULL)) throw std::runtime_error("waveOutOpen failed for the original audio format");
    }
    void reset() {
        if(!wave) return; waveOutReset(wave);
        for(auto& s:slots) if(s.prepared) { waveOutUnprepareHeader(wave,&s.h,sizeof(s.h)); s.prepared=false; }
    }
    ~Output() { reset(); if(wave) waveOutClose(wave); }
    Slot* freeSlot() {
        for(auto& s:slots) if(!s.prepared || (s.h.dwFlags&WHDR_DONE)) {
            if(s.prepared) waveOutUnprepareHeader(wave,&s.h,sizeof(s.h));
            s.prepared=false; return &s;
        }
        return nullptr;
    }
    bool pending() { for(auto& s:slots) if(s.prepared && !(s.h.dwFlags&WHDR_DONE)) return true; return false; }
    void submit(Slot& s,Bytes b) {
        s.data=std::move(b); s.h={}; s.h.lpData=(char*)s.data.data(); s.h.dwBufferLength=(DWORD)s.data.size();
        if(waveOutPrepareHeader(wave,&s.h,sizeof(s.h))) throw std::runtime_error("Prepare output failed");
        s.prepared=true;
        if(waveOutWrite(wave,&s.h,sizeof(s.h))) throw std::runtime_error("Output failed");
    }
    long long played() { MMTIME t{}; t.wType=TIME_SAMPLES; waveOutGetPosition(wave,&t,sizeof(t)); return t.wType==TIME_SAMPLES?t.u.sample:0; }
};
