#define REVERSE_BUILD
#include "api.h"
#include "engine.h"

namespace {
thread_local wchar_t lastError[512]{};
void copyError(wchar_t* target,const char* message) {
    MultiByteToWideChar(CP_UTF8,0,message,-1,target,512); target[511]=0;
}
struct Guard {
    CRITICAL_SECTION* p;
    Guard(CRITICAL_SECTION& cs):p(&cs) { EnterCriticalSection(p); }
    ~Guard() { LeaveCriticalSection(p); }
};
enum { Play=1,Pause,Stop,Seek,Volume,Export,Quit };
struct Session {
    CRITICAL_SECTION api{},data{};
    HANDLE ready=nullptr,wake=nullptr,ack=nullptr,thread=nullptr;
    std::wstring path,exportPath;
    int mode=2,command=0,state=0,volume=100;
    double chunk=2,argument=0,length=0,position=0;
    DWORD hz=0,channels=0;
    HRESULT result=E_PENDING;
    wchar_t error[512]{};
    Session() {
        InitializeCriticalSection(&api); InitializeCriticalSection(&data);
        ready=CreateEventW(nullptr,TRUE,FALSE,nullptr);
        wake=CreateEventW(nullptr,TRUE,FALSE,nullptr);
        ack=CreateEventW(nullptr,FALSE,FALSE,nullptr);
    }
    ~Session() {
        if(thread) CloseHandle(thread);
        if(ready) CloseHandle(ready); if(wake) CloseHandle(wake); if(ack) CloseHandle(ack);
        DeleteCriticalSection(&api); DeleteCriticalSection(&data);
    }
};
DWORD WINAPI worker(void* value) {
    auto& s=*static_cast<Session*>(value);
    bool com=false,mf=false,initialized=false;
    try {
        check(CoInitializeEx(nullptr,COINIT_MULTITHREADED),"Worker COM"); com=true;
        check(MFStartup(MF_VERSION),"Media Foundation startup"); mf=true;
        {
            Reader reader; reader.open(s.path.c_str());
            const auto fmt=reader.fmt;
            IndexedMP3 index;
            Bytes cache;
            const bool backward=s.mode!=0, blocks=s.mode==2, indexed=blocks && reader.mp3Quantum;
            if(indexed) { index.open(s.path.c_str(),reader.initialGap,fmt); reader.total=index.total; }
            if(!blocks) cache=reader.all();
            const long long total=reader.total,step=std::max(1LL,llround(s.chunk*fmt.hz));
            if(total<=0) throw std::runtime_error("No audio samples");
            long long cursor=backward?total:0,anchor=cursor;
            Output output;
            Bytes prepared;
            long long preparedLo=-1,preparedHi=-1;
            auto current=[&]() {
                return std::clamp(anchor+(backward?-1:1)*(output.wave?output.played():0LL),0LL,total);
            };
            auto render=[&](long long lo,long long hi) {
                Bytes b=blocks?(indexed?index.range(lo,hi):reader.range(lo,hi)):
                    Bytes(cache.begin()+lo*fmt.align(),cache.begin()+hi*fmt.align());
                if(backward) reversePCM(b,fmt.align());
                return b;
            };
            auto publish=[&]() {
                Guard g(s.data); s.position=double(current())/fmt.hz;
            };
            {
                Guard g(s.data);
                s.hz=fmt.hz;s.channels=fmt.channels;s.length=double(total)/fmt.hz;
                s.position=double(cursor)/fmt.hz;s.state=0;s.result=S_OK;
            }
            initialized=true; SetEvent(s.ready);
            bool quit=false; int state=0,level=100;
            while(!quit) {
                if(WaitForSingleObject(s.wake,0)==WAIT_OBJECT_0) {
                    ResetEvent(s.wake);
                    int cmd; double arg; std::wstring exportPath;
                    { Guard g(s.data);cmd=s.command;arg=s.argument;exportPath=s.exportPath; }
                    HRESULT hr=S_OK;
                    try {
                        if(cmd==Quit) quit=true;
                        else if(cmd==Play) {
                            if(!output.wave) output.open(fmt);
                            if(state==3 || (backward?current()==0:current()==total)) {
                                output.reset();cursor=anchor=backward?total:0;
                            }
                            waveOutRestart(output.wave);
                            state=1;
                        } else if(cmd==Pause) {
                            if(state==1) { waveOutPause(output.wave);state=2; }
                        } else if(cmd==Seek || cmd==Stop || cmd==Volume) {
                            long long target=current();
                            if(cmd==Seek) target=std::clamp(llround(arg*fmt.hz),0LL,total);
                            if(cmd==Seek) {
                                // Validate the new source range before acknowledging a seek.
                                // Clients can keep the old backend until this succeeds.
                                long long lo=backward?std::max(0LL,target-step):target;
                                long long hi=backward?target:std::min(total,target+step);
                                Bytes next;
                                if(hi>lo) next=render(lo,hi);
                                prepared=std::move(next);preparedLo=lo;preparedHi=hi;
                            }
                            if(cmd==Stop) { target=backward?total:0;state=0; }
                            if(cmd==Volume) level=(int)arg;
                            if(output.wave) output.reset();
                            cursor=anchor=target;
                            if(state==3 && cmd==Seek) state=2;
                        } else if(cmd==Export) {
                            if(state==1 || state==2) throw std::runtime_error("Stop playback before export");
                            Wav wav; wav.open(exportPath.c_str(),fmt);
                            long long pos=arg<0?(backward?total:0):std::clamp(llround(arg*fmt.hz),0LL,total);
                            while(backward?pos>0:pos<total) {
                                long long lo=backward?std::max(0LL,pos-step):pos;
                                long long hi=backward?pos:std::min(total,pos+step);
                                wav.write(render(lo,hi));pos=backward?lo:hi;
                            }
                        }
                    } catch(const std::exception& e) {
                        hr=E_FAIL;Guard g(s.data);copyError(s.error,e.what());
                    }
                    {
                        Guard g(s.data);s.state=state;s.volume=level;s.result=hr;
                        s.position=double(current())/fmt.hz;
                    }
                    SetEvent(s.ack);
                    if(quit) break;
                }
                if(state==1) {
                    const bool end=backward?cursor<=0:cursor>=total;
                    if(end && !output.pending()) {
                        state=3;anchor=backward?0:total;output.reset();cursor=anchor;
                    } else if(!end) {
                        auto slot=output.freeSlot();
                        if(slot) {
                            long long lo=backward?std::max(0LL,cursor-step):cursor;
                            long long hi=backward?cursor:std::min(total,cursor+step);
                            Bytes b;
                            if(lo==preparedLo && hi==preparedHi && !prepared.empty()) {
                                b=std::move(prepared);preparedLo=preparedHi=-1;
                            } else b=render(lo,hi);
                            if(level!=100) for(size_t i=0;i+1<b.size();i+=2) {
                                short sample;memcpy(&sample,b.data()+i,2);
                                sample=(short)(int(sample)*level/100);memcpy(b.data()+i,&sample,2);
                            }
                            output.submit(*slot,std::move(b));cursor=backward?lo:hi;
                        }
                    }
                }
                publish();{Guard g(s.data);s.state=state;}
                // Manual-reset wake stays signalled until the command is taken.
                // Paused/stopped sessions consume no polling CPU.
                WaitForSingleObject(s.wake,state==1?2:INFINITE);
            }
        }
    } catch(const std::exception& e) {
        Guard g(s.data);copyError(s.error,e.what());s.result=E_FAIL;s.state=4;
    } catch(...) {
        Guard g(s.data);lstrcpyW(s.error,L"Unexpected reverse worker failure");s.result=E_FAIL;s.state=4;
    }
    if(mf) MFShutdown();if(com) CoUninitialize();
    if(!initialized) SetEvent(s.ready);
    SetEvent(s.ack);
    return 0;
}
HRESULT request(Session* s,int cmd,double arg=0,const wchar_t* path=nullptr) {
    if(!s) return E_POINTER;
    Guard api(s->api);
    if(WaitForSingleObject(s->thread,0)==WAIT_OBJECT_0) return E_FAIL;
    ResetEvent(s->ack);
    {Guard g(s->data);s->command=cmd;s->argument=arg;if(path)s->exportPath=path;}
    SetEvent(s->wake);
    HANDLE handles[]={s->ack,s->thread};
    DWORD wait=WaitForMultipleObjects(2,handles,FALSE,INFINITE);
    Guard g(s->data);
    return wait==WAIT_OBJECT_0?s->result:E_FAIL;
}
}
REV DWORD WINAPI Reverse_Version() {return 0x10000;}
REV HRESULT WINAPI Reverse_OpenW(const wchar_t* file,int mode,double chunk,void** out) {
    if(!out) return E_POINTER;*out=nullptr;
    if(!file || mode<0 || mode>2 || !std::isfinite(chunk) || chunk<0.05 || chunk>30) return E_INVALIDARG;
    if(!wcsncmp(file,L"http:",5)||!wcsncmp(file,L"https:",6)) return E_INVALIDARG;
    try {
        auto s=new Session();s->path=file;s->mode=mode;s->chunk=chunk;
        if(!s->ready||!s->wake||!s->ack) {delete s;return E_OUTOFMEMORY;}
        s->thread=CreateThread(nullptr,0,worker,s,0,nullptr);
        if(!s->thread) {delete s;return HRESULT_FROM_WIN32(GetLastError());}
        WaitForSingleObject(s->ready,INFINITE);
        if(FAILED(s->result)) {
            lstrcpynW(lastError,s->error,512);WaitForSingleObject(s->thread,INFINITE);
            HRESULT hr=s->result;delete s;return hr;
        }
        *out=s;return S_OK;
    } catch(...) { lstrcpyW(lastError,L"Not enough memory to open reverse session");return E_OUTOFMEMORY; }
}
REV void WINAPI Reverse_Close(void* p) {
    auto s=static_cast<Session*>(p);if(!s)return;
    request(s,Quit);WaitForSingleObject(s->thread,INFINITE);delete s;
}
REV HRESULT WINAPI Reverse_Play(void* p) {return request(static_cast<Session*>(p),Play);}
REV HRESULT WINAPI Reverse_Pause(void* p) {return request(static_cast<Session*>(p),Pause);}
REV HRESULT WINAPI Reverse_Stop(void* p) {return request(static_cast<Session*>(p),Stop);}
REV HRESULT WINAPI Reverse_Seek(void* p,double v) {return std::isfinite(v)?request(static_cast<Session*>(p),Seek,v):E_INVALIDARG;}
REV HRESULT WINAPI Reverse_Volume(void* p,int v) {return v>=0&&v<=100?request(static_cast<Session*>(p),Volume,v):E_INVALIDARG;}
REV double WINAPI Reverse_Duration(void* p) {auto s=static_cast<Session*>(p);if(!s)return 0;Guard g(s->data);return s->length;}
REV double WINAPI Reverse_Position(void* p) {auto s=static_cast<Session*>(p);if(!s)return 0;Guard g(s->data);return s->position;}
REV int WINAPI Reverse_State(void* p) {auto s=static_cast<Session*>(p);if(!s)return 0;Guard g(s->data);return s->state;}
REV HRESULT WINAPI Reverse_Format(void* p,DWORD* hz,DWORD* channels) {
    auto s=static_cast<Session*>(p);if(!s||!hz||!channels)return E_POINTER;
    Guard g(s->data);*hz=s->hz;*channels=s->channels;return S_OK;
}
REV void WINAPI Reverse_ErrorW(void* p,wchar_t* out,int capacity) {
    if(!out||capacity<=0)return;
    auto s=static_cast<Session*>(p);
    if(s){Guard g(s->data);lstrcpynW(out,s->error,capacity);}else lstrcpynW(out,lastError,capacity);
}
REV HRESULT WINAPI Reverse_ExportW(void* p,const wchar_t* path,double start) {
    if(!path||!std::isfinite(start))return E_INVALIDARG;
    if(GetFileAttributesW(path)!=INVALID_FILE_ATTRIBUTES)return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    return request(static_cast<Session*>(p),Export,start,path);
}
