#pragma once
#include <strsafe.h>


//void debug(_In_ _Printf_format_string_ LPCWSTR, ...)
//{
//    __VA_ARG
//}
#define debug(fmt, ...)     do {                                \
    WCHAR buf[256] = { 0 };                                     \
    StringCchPrintfW(buf, _countof(buf), L##fmt, __VA_ARGS__);  \
    StringCchCatW(buf, _countof(buf), L"\n");                   \
    OutputDebugStringW(buf);                                    \
} while (false)
