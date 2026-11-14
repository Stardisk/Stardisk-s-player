#include "base.h"
// Compiler-generated aggregate copies are implemented locally, not supplied
// by a CRT. Volatile byte loops keep the optimizer from calling these again.
extern "C" void* memset(void* destination, int value, SIZE_T count)
{
    auto p = static_cast<volatile unsigned char*>(destination);
    while (count--)
        *p++ = (unsigned char)value;
    return destination;
}
extern "C" void* memcpy(void* destination, const void* source, SIZE_T count)
{
    auto p = static_cast<volatile unsigned char*>(destination);
    auto q = static_cast<const volatile unsigned char*>(source);
    while (count--)
        *p++ = *q++;
    return destination;
}
extern "C" void* memmove(void* destination, const void* source, SIZE_T count)
{
    auto p = static_cast<volatile unsigned char*>(destination);
    auto q = static_cast<const volatile unsigned char*>(source);
    if (p < q)
        while (count--)
            *p++ = *q++;
    else
    {
        p += count;
        q += count;
        while (count--)
            *--p = *--q;
    }
    return destination;
}
extern "C" int memcmp(const void* first, const void* second, SIZE_T count)
{
    auto p = static_cast<const unsigned char*>(first);
    auto q = static_cast<const unsigned char*>(second);
    while (count--)
    {
        if (*p != *q)
            return *p - *q;
        ++p;
        ++q;
    }
    return 0;
}
void operator delete(void* p) noexcept
{
    freeMemory(p);
}
void operator delete(void* p, SIZE_T) noexcept
{
    freeMemory(p);
}
extern "C" int _purecall()
{
    ExitProcess(ERROR_INVALID_FUNCTION);
}
extern "C" void __cxa_pure_virtual()
{
    ExitProcess(ERROR_INVALID_FUNCTION);
}
