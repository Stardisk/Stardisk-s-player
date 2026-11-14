#include "base.h"

Text encodeStream(const Text& value)
{
    Text result;
    for (int i = 0; i < value.size(); ++i)
    {
        wchar_t c = value.c_str()[i];
        if (c == L'\\')
            result += L"\\\\";
        else if (c < 32 || c > 126)
        {
            wchar_t b[7] = {L'\\', L'u', 0, 0, 0, 0, 0};
            for (int k = 0; k < 4; ++k)
                b[5 - k] = L"0123456789ABCDEF"[(c >> (k * 4)) & 15];
            result += b;
        }
        else
        {
            wchar_t b[2] = {c, 0};
            result += b;
        }
    }
    return result;
}

Text decodeStream(const Text& value)
{
    Text result;
    for (int i = 0; i < value.size(); ++i)
    {
        wchar_t c = value.c_str()[i];
        if (c == L'\\' && i + 1 < value.size())
        {
            if (value.c_str()[i + 1] == L'\\')
            {
                c = L'\\';
                ++i;
            }
            else if (value.c_str()[i + 1] == L'u' && i + 5 < value.size())
            {
                int n = 0;
                bool ok = true;
                for (int k = 2; k < 6; ++k)
                {
                    wchar_t d = value.c_str()[i + k];
                    int v = d >= L'0' && d <= L'9'   ? d - L'0'
                            : d >= L'A' && d <= L'F' ? d - L'A' + 10
                            : d >= L'a' && d <= L'f' ? d - L'a' + 10
                                                     : -1;
                    if (v < 0)
                    {
                        ok = false;
                        break;
                    }
                    n = n * 16 + v;
                }
                if (ok)
                {
                    c = wchar_t(n);
                    i += 5;
                }
            }
        }
        wchar_t b[2] = {c, 0};
        result += b;
    }
    return result;
}
