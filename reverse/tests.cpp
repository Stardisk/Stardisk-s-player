#define NOMINMAX
#include "api.h"
#include <cstdio>
#include <cmath>

int wmain(int argc,wchar_t** argv) {
    ReverseApi api;
    if(argc==2 && !lstrcmpW(argv[1],L"--missing")) {
        bool missing=!api.load();printf("%s missing DLL does not prevent startup\n",missing?"PASS":"FAIL");return missing?0:1;
    }
    if(argc!=2||!api.load())return 100;
    int failures=0;auto test=[&](bool ok,const char* label){printf("%s %s\n",ok?"PASS":"FAIL",label);if(!ok)++failures;};
    void* s=nullptr;
    test(api.version()==0x10000,"native x64 ABI version");
    test(FAILED(api.open(L"Z:\\missing-reverse-test.wav",2,.13,&s)) && !s,"open error returns no dangling session");
    test(SUCCEEDED(api.open(argv[1],2,.13,&s)) && s,"open Unicode local file");
    if(!s)return 1;
    DWORD hz=0,channels=0;api.format(s,&hz,&channels);
    test(hz==44100&&channels==1,"source sample rate and mono preserved");
    double duration=api.duration(s);
    test(duration>29.9&&duration<30.1,"correct duration");
    test(SUCCEEDED(api.volume(s,0)),"set mute before output starts");
    test(SUCCEEDED(api.seek(s,3)),"seek while stopped");
    test(SUCCEEDED(api.play(s)),"start reverse worker output");
    Sleep(220);double position=api.position(s);
    test(position<2.95&&position>2.5,"reverse clock decreases");
    test(SUCCEEDED(api.pause(s)),"pause");
    position=api.position(s);Sleep(140);
    test(std::abs(position-api.position(s))<.015,"pause holds position");
    test(SUCCEEDED(api.seek(s,1.5)),"seek while paused");
    test(std::abs(api.position(s)-1.5)<.001&&api.state(s)==2,"seek preserves pause");
    test(SUCCEEDED(api.play(s)),"resume after paused seek");
    test(SUCCEEDED(api.volume(s,0)),"volume change rebuilds queued sound at current position");
    test(SUCCEEDED(api.stop(s)),"stop");
    test(std::abs(api.position(s)-duration)<.001&&api.state(s)==0,"reverse stop resets to end");
    for(int i=0;i<2;i++) {
        test(SUCCEEDED(api.seek(s,.12))&&SUCCEEDED(api.play(s)),"play short tail toward beginning");
        DWORD started=GetTickCount();while(api.state(s)==1 && GetTickCount()-started<2000)Sleep(5);
        test(api.state(s)==3&&api.position(s)==0,"natural reverse completion at zero; restart works");
        api.volume(s,0);
        test(api.state(s)==3,"volume at natural end preserves completion notification");
    }
    api.close(s);s=nullptr;api.unload();
    HANDLE file=CreateFileW(argv[1],GENERIC_READ|GENERIC_WRITE,0,nullptr,OPEN_EXISTING,0,nullptr);
    test(file!=INVALID_HANDLE_VALUE,"close releases file and worker before DLL unload");
    if(file!=INVALID_HANDLE_VALUE)CloseHandle(file);
    printf("Failures: %d\n",failures);return failures?1:0;
}
