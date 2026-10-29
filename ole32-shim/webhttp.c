/*
 * Windows.Web.Http for the ole32 shim, on top of WinHTTP.
 *
 * Wine has no Windows.Web.Http. Office's React Native surfaces (react-native-win32: OneNote's
 * navigation pane, parts of Outlook) do all their fetch() and XMLHttpRequest traffic through it: an
 * HttpClient over react-native-windows' own redirect filter over an HttpBaseProtocolFilter,
 * HttpRequestMessage with HttpStringContent or HttpBufferContent, and the response body read back
 * through a Windows.Storage.Streams.DataReader, which Wine lacks as well. This file implements that
 * much of the API. A request runs synchronously on WinHTTP in a thread-pool thread and completes
 * the IAsyncOperationWithProgress the caller is waiting on; the response body is read whole, so
 * every other operation completes at once. Members nothing has been seen calling answer E_NOTIMPL
 * and say so in the debug log ("ms365 webhttp: ... not implemented").
 *
 * Layouts follow Windows.Foundation.UniversalApiContract.winmd. Parameterized interface IIDs are the
 * WinRT pinterface hashes (most of them also appear in react-native-win32.dll). Every out parameter
 * gets its exact width: a WinRT boolean is one byte, an enum four. The classes are served from the
 * shim's DllGetActivationFactory (the launcher registers them for ms365shim.dll) and from its
 * RoGetActivationFactory fallback; both come here through webhttp_get_factory.
 */
#include <windows.h>
#include <winhttp.h>
#include <stdarg.h>
#include <string.h>
#include <wchar.h>

typedef void *HSTRING;
typedef unsigned char boolean8;   /* WinRT boolean */

HRESULT webhttp_get_factory(const WCHAR *name, REFIID iid, void **out);

#define E_BOUNDS_                     ((HRESULT)0x8000000B)
#define E_ILLEGAL_METHOD_CALL_        ((HRESULT)0x8000000E)
#define E_ILLEGAL_DELEGATE_ASSIGNMENT_ ((HRESULT)0x80000018)

static const GUID IID_Inspectable      = { 0xaf86e2e0, 0xb12d, 0x4c6a, { 0x9c, 0x5a, 0xd7, 0xaa, 0x65, 0x10, 0x1e, 0x90 } };
static const GUID IID_AgileObject      = { 0x94ea2b94, 0xe9cc, 0x49e0, { 0xc0, 0xff, 0xee, 0x64, 0xca, 0x8f, 0x5b, 0x90 } };
static const GUID IID_ActivationFactory= { 0x00000035, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
static const GUID IID_BufferByteAccess = { 0x905a0fef, 0xbc53, 0x11df, { 0x8c, 0x49, 0x00, 0x1e, 0x4f, 0xc6, 0x86, 0xda } };
/* private: QueryInterface hands out the object behind any of our interface pointers */
static const GUID IID_Self             = { 0x6d5b8f2e, 0x4c1a, 0x4e7b, { 0x9a, 0x3d, 0x0b, 0x8c, 0x1f, 0x2e, 0x3d, 0x4a } };

static const GUID IID_AsyncInfo                  = { 0x00000036, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
static const GUID IID_Closable                   = { 0x30d5a829, 0x7fa4, 0x4026, { 0x83, 0xbb, 0xd7, 0x5b, 0xae, 0x4e, 0xa9, 0x9e } };
static const GUID IID_Stringable                 = { 0x96369f54, 0x8eb6, 0x48f0, { 0xab, 0xce, 0xc1, 0xb2, 0x11, 0xe6, 0x27, 0xc3 } };
static const GUID IID_UriFactory                 = { 0x44a9796f, 0x723e, 0x4fdf, { 0xa2, 0x18, 0x03, 0x3e, 0x75, 0xb0, 0xc0, 0x84 } };
static const GUID IID_Buffer                     = { 0x905a0fe0, 0xbc53, 0x11df, { 0x8c, 0x49, 0x00, 0x1e, 0x4f, 0xc6, 0x86, 0xda } };
static const GUID IID_InputStream                = { 0x905a0fe2, 0xbc53, 0x11df, { 0x8c, 0x49, 0x00, 0x1e, 0x4f, 0xc6, 0x86, 0xda } };
static const GUID IID_DataReader                 = { 0xe2b50029, 0xb4c1, 0x4314, { 0xa4, 0xb8, 0xfb, 0x81, 0x3a, 0x2f, 0x27, 0x5e } };
static const GUID IID_DataReaderFactory          = { 0xd7527847, 0x57da, 0x4e15, { 0x91, 0x4c, 0x06, 0x80, 0x66, 0x99, 0xa0, 0x98 } };
static const GUID IID_DataReaderStatics          = { 0x11fcbfc8, 0xf93a, 0x471b, { 0xb1, 0x21, 0xf3, 0x79, 0xe3, 0x49, 0x31, 0x3c } };
static const GUID IID_Filter                     = { 0x71c89b09, 0xe131, 0x4b54, { 0xa5, 0x3c, 0xeb, 0x43, 0xff, 0x37, 0xe9, 0xbb } };
static const GUID IID_Filter2                    = { 0x2ec30013, 0x9427, 0x4900, { 0xa0, 0x17, 0xfa, 0x7d, 0xa3, 0xb5, 0xc9, 0xae } };
static const GUID IID_Filter3                    = { 0xd43f4d4c, 0xbd42, 0x43ae, { 0x87, 0x17, 0xad, 0x2c, 0x8f, 0x4b, 0x29, 0x37 } };
static const GUID IID_Filter4                    = { 0x9fe36ccf, 0x2983, 0x4893, { 0x94, 0x1f, 0xeb, 0x51, 0x8c, 0xa8, 0xce, 0xf9 } };
static const GUID IID_Filter5                    = { 0x416e4993, 0x31e3, 0x4816, { 0xbf, 0x09, 0xe0, 0x18, 0xee, 0x8d, 0xc1, 0xf5 } };
static const GUID IID_HttpFilter                 = { 0xa4cb6dd5, 0x0902, 0x439e, { 0xbf, 0xd7, 0xe1, 0x25, 0x52, 0xb1, 0x65, 0xce } };
static const GUID IID_CacheControl               = { 0xc77e1cb4, 0x3cea, 0x4eb5, { 0xac, 0x85, 0x04, 0xe1, 0x86, 0xe6, 0x3a, 0xb7 } };
static const GUID IID_Client                     = { 0x7fda1151, 0x3574, 0x4880, { 0xa8, 0xba, 0xe6, 0xb1, 0xe0, 0x06, 0x1f, 0x3d } };
static const GUID IID_Client2                    = { 0xcdd83348, 0xe8b7, 0x4cec, { 0xb1, 0xb0, 0xdc, 0x45, 0x5f, 0xe7, 0x2c, 0x92 } };
static const GUID IID_Client3                    = { 0x1172fd01, 0x9899, 0x4194, { 0x96, 0x3f, 0x8f, 0x9d, 0x72, 0xa7, 0xec, 0x15 } };
static const GUID IID_ClientFactory              = { 0xc30c4eca, 0xe3fa, 0x4f99, { 0xaf, 0xb4, 0x63, 0xcc, 0x65, 0x00, 0x94, 0x62 } };
static const GUID IID_Method                     = { 0x728d4022, 0x700d, 0x4fe0, { 0xaf, 0xa5, 0x40, 0x29, 0x9c, 0x58, 0xdb, 0xfd } };
static const GUID IID_MethodFactory              = { 0x3c51d10d, 0x36d7, 0x40f8, { 0xa8, 0x6d, 0xe7, 0x59, 0xca, 0xf2, 0xf8, 0x3f } };
static const GUID IID_MethodStatics              = { 0x64d171f0, 0xd99a, 0x4153, { 0x8d, 0xc6, 0xd6, 0x8c, 0xc4, 0xcc, 0xe3, 0x17 } };
static const GUID IID_Request                    = { 0xf5762b3c, 0x74d4, 0x4811, { 0xb5, 0xdc, 0x9f, 0x8b, 0x4e, 0x2f, 0x9a, 0xbf } };
static const GUID IID_Request2                   = { 0xc3c60489, 0x62c2, 0x4a3f, { 0x95, 0x54, 0x22, 0x6e, 0x7c, 0x60, 0xbd, 0x96 } };
static const GUID IID_RequestFactory             = { 0x5bac994e, 0x3886, 0x412e, { 0xae, 0xc3, 0x52, 0xec, 0x7f, 0x25, 0x61, 0x6f } };
static const GUID IID_Response                   = { 0xfee200fb, 0x8664, 0x44e0, { 0x95, 0xd9, 0x42, 0x69, 0x61, 0x99, 0xbf, 0xfc } };
static const GUID IID_ResponseFactory            = { 0x52a8af99, 0xf095, 0x43da, { 0xb6, 0x0f, 0x7c, 0xfc, 0x2b, 0xc7, 0xea, 0x2f } };
static const GUID IID_Content                    = { 0x6b14a441, 0xfba7, 0x4bd2, { 0xaf, 0x0a, 0x83, 0x9d, 0xe7, 0xc2, 0x95, 0xda } };
static const GUID IID_StringContentFactory       = { 0x46649d5b, 0x2e93, 0x48eb, { 0x8e, 0x61, 0x19, 0x67, 0x78, 0x78, 0xe5, 0x7f } };
static const GUID IID_BufferContentFactory       = { 0xbc20c193, 0xc41f, 0x4ff7, { 0x91, 0x23, 0x64, 0x35, 0x73, 0x6e, 0xad, 0xc2 } };
static const GUID IID_RequestHeaders             = { 0xaf40329b, 0xb544, 0x469b, { 0x86, 0xb9, 0xac, 0x3d, 0x46, 0x6f, 0xea, 0x36 } };
static const GUID IID_ResponseHeaders            = { 0x7a990969, 0xfa3f, 0x41ed, { 0xaa, 0xc6, 0xbf, 0x95, 0x79, 0x75, 0xc1, 0x6b } };
static const GUID IID_ContentHeaders             = { 0x40612a44, 0x47ae, 0x4b7e, { 0x91, 0x24, 0x69, 0x62, 0x8b, 0x64, 0xaa, 0x18 } };
static const GUID IID_MediaType                  = { 0x16b28533, 0xe728, 0x4fcb, { 0xbd, 0xb0, 0x08, 0xa4, 0x31, 0xa1, 0x48, 0x44 } };
static const GUID IID_MediaTypeFactory           = { 0xbed747a8, 0xcd17, 0x42dd, { 0x93, 0x67, 0xab, 0x9c, 0x5b, 0x56, 0xdd, 0x7d } };
static const GUID IID_MediaTypeStatics           = { 0xe04d83df, 0x1d41, 0x4d8c, { 0xa2, 0xde, 0x6f, 0xd2, 0xed, 0x87, 0x39, 0x9b } };
static const GUID IID_ContentCodings             = { 0x7d221721, 0xa6db, 0x436e, { 0x8e, 0x83, 0x91, 0x59, 0x61, 0x92, 0x81, 0x9c } };
/* parameterized: IAsyncOperationWithProgress<HttpResponseMessage, HttpProgress> and friends */
static const GUID IID_OpResponse      = { 0x5d144364, 0x77d7, 0x5eca, { 0x8b, 0x09, 0x93, 0x6a, 0x69, 0x44, 0x66, 0x52 } };
static const GUID IID_OpInputStream64 = { 0x455aa601, 0xf13e, 0x5dee, { 0xb9, 0xcb, 0x16, 0xb5, 0x31, 0x99, 0x63, 0x27 } };
static const GUID IID_OpBuffer64      = { 0xad960e7f, 0xd73b, 0x56e4, { 0xa5, 0x8c, 0x6e, 0xc7, 0x67, 0x8c, 0xfd, 0x88 } };
static const GUID IID_OpString64      = { 0xc8bbcb29, 0x6b64, 0x5ce2, { 0xa8, 0x31, 0x03, 0x8f, 0x6e, 0x02, 0x19, 0x9e } };
static const GUID IID_OpU64U64        = { 0x8f1db6e3, 0x6556, 0x5516, { 0x82, 0x5c, 0x10, 0x21, 0xee, 0x27, 0xcd, 0x0c } };
static const GUID IID_OpBuffer32      = { 0xd26b2819, 0x897f, 0x5c7d, { 0x84, 0xd6, 0x56, 0xd7, 0x96, 0x56, 0x14, 0x31 } };
static const GUID IID_OpU32           = { 0xef60385f, 0xbe78, 0x584b, { 0xaa, 0xef, 0x78, 0x29, 0xad, 0xa2, 0xb0, 0xde } };
static const GUID IID_RefU64          = { 0x6755e376, 0x53bb, 0x568b, { 0xa1, 0x1d, 0x17, 0x23, 0x98, 0x68, 0x30, 0x9e } };
static const GUID IID_MapSO           = { 0x1b0d3570, 0x0877, 0x5ec2, { 0x8a, 0x2c, 0x3b, 0x95, 0x39, 0x50, 0x6a, 0xca } };
static const GUID IID_MapViewSO       = { 0xbb78502a, 0xf79d, 0x54fa, { 0x92, 0xc9, 0x90, 0xc5, 0x03, 0x9f, 0xdf, 0x7e } };
static const GUID IID_KvpSO           = { 0x09335560, 0x6c6b, 0x5a26, { 0x93, 0x48, 0x97, 0xb7, 0x81, 0x13, 0x2b, 0x20 } };
static const GUID IID_IterableSO      = { 0xfe2f3d47, 0x5d47, 0x5499, { 0x83, 0x74, 0x43, 0x0c, 0x7c, 0xda, 0x02, 0x04 } };
static const GUID IID_IteratorSO      = { 0x5db5fa32, 0x707c, 0x5849, { 0xa0, 0x6b, 0x91, 0xc8, 0xeb, 0x9d, 0x10, 0xe8 } };
static const GUID IID_MapSS           = { 0xf6d1f700, 0x49c2, 0x52ae, { 0x81, 0x54, 0x82, 0x6f, 0x99, 0x08, 0x77, 0x3c } };
static const GUID IID_MapViewSS       = { 0xac7f26f2, 0xfeb7, 0x5b2a, { 0x8a, 0xc4, 0x34, 0x5b, 0xc6, 0x2c, 0xae, 0xde } };
static const GUID IID_KvpSS           = { 0x60310303, 0x49c5, 0x52e6, { 0xab, 0xc6, 0xa9, 0xb3, 0x6e, 0xcc, 0xc7, 0x16 } };
static const GUID IID_IterableSS      = { 0xe9bdaaf0, 0xcbf6, 0x5c72, { 0xbe, 0x90, 0x29, 0xcb, 0xf3, 0xa1, 0x31, 0x9b } };
static const GUID IID_IteratorSS      = { 0x05eb86f1, 0x7140, 0x5517, { 0xb8, 0x8d, 0xcb, 0xae, 0xbe, 0x57, 0xe6, 0xb1 } };

/* ---- runtime entry points (combase, winhttp), resolved on first use, never from DllMain ---- */
static HRESULT (WINAPI *p_WindowsCreateString)(const WCHAR *, UINT32, HSTRING *);
static const WCHAR *(WINAPI *p_WindowsGetStringRawBuffer)(HSTRING, UINT32 *);
static HRESULT (WINAPI *p_RoGetActivationFactory)(HSTRING, REFIID, void **);
static HRESULT (WINAPI *p_WindowsDeleteString)(HSTRING);
static HINTERNET (WINAPI *p_WinHttpOpen)(LPCWSTR, DWORD, LPCWSTR, LPCWSTR, DWORD);
static HINTERNET (WINAPI *p_WinHttpConnect)(HINTERNET, LPCWSTR, INTERNET_PORT, DWORD);
static HINTERNET (WINAPI *p_WinHttpOpenRequest)(HINTERNET, LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR *, DWORD);
static BOOL (WINAPI *p_WinHttpAddRequestHeaders)(HINTERNET, LPCWSTR, DWORD, DWORD);
static BOOL (WINAPI *p_WinHttpSendRequest)(HINTERNET, LPCWSTR, DWORD, LPVOID, DWORD, DWORD, DWORD_PTR);
static BOOL (WINAPI *p_WinHttpReceiveResponse)(HINTERNET, LPVOID);
static BOOL (WINAPI *p_WinHttpQueryHeaders)(HINTERNET, DWORD, LPCWSTR, LPVOID, LPDWORD, LPDWORD);
static BOOL (WINAPI *p_WinHttpQueryDataAvailable)(HINTERNET, LPDWORD);
static BOOL (WINAPI *p_WinHttpReadData)(HINTERNET, LPVOID, DWORD, LPDWORD);
static BOOL (WINAPI *p_WinHttpSetOption)(HINTERNET, DWORD, LPVOID, DWORD);
static BOOL (WINAPI *p_WinHttpCloseHandle)(HINTERNET);
static BOOL (WINAPI *p_WinHttpCrackUrl)(LPCWSTR, DWORD, DWORD, LPURL_COMPONENTS);
static HINTERNET g_session;

static INIT_ONCE g_init_once = INIT_ONCE_STATIC_INIT;
static BOOL CALLBACK init_once(INIT_ONCE *once, void *param, void **ctx)
{
    (void)once; (void)param; (void)ctx;
    HMODULE cb = GetModuleHandleA("combase.dll");
    if (!cb) cb = LoadLibraryA("combase.dll");
    if (cb) {
        p_WindowsCreateString = (void *)GetProcAddress(cb, "WindowsCreateString");
        p_WindowsGetStringRawBuffer = (void *)GetProcAddress(cb, "WindowsGetStringRawBuffer");
        p_WindowsDeleteString = (void *)GetProcAddress(cb, "WindowsDeleteString");
        p_RoGetActivationFactory = (void *)GetProcAddress(cb, "RoGetActivationFactory");
    }
    HMODULE wh = LoadLibraryA("winhttp.dll");
    if (wh) {
        p_WinHttpOpen = (void *)GetProcAddress(wh, "WinHttpOpen");
        p_WinHttpConnect = (void *)GetProcAddress(wh, "WinHttpConnect");
        p_WinHttpOpenRequest = (void *)GetProcAddress(wh, "WinHttpOpenRequest");
        p_WinHttpAddRequestHeaders = (void *)GetProcAddress(wh, "WinHttpAddRequestHeaders");
        p_WinHttpSendRequest = (void *)GetProcAddress(wh, "WinHttpSendRequest");
        p_WinHttpReceiveResponse = (void *)GetProcAddress(wh, "WinHttpReceiveResponse");
        p_WinHttpQueryHeaders = (void *)GetProcAddress(wh, "WinHttpQueryHeaders");
        p_WinHttpQueryDataAvailable = (void *)GetProcAddress(wh, "WinHttpQueryDataAvailable");
        p_WinHttpReadData = (void *)GetProcAddress(wh, "WinHttpReadData");
        p_WinHttpSetOption = (void *)GetProcAddress(wh, "WinHttpSetOption");
        p_WinHttpCloseHandle = (void *)GetProcAddress(wh, "WinHttpCloseHandle");
        p_WinHttpCrackUrl = (void *)GetProcAddress(wh, "WinHttpCrackUrl");
        if (p_WinHttpOpen)
            g_session = p_WinHttpOpen(L"Mozilla/5.0 (Windows NT 10.0; Win64; x64)", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                      WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    }
    return TRUE;
}
static BOOL init(void)
{
    InitOnceExecuteOnce(&g_init_once, init_once, NULL, NULL);
    return p_WindowsCreateString && p_WindowsGetStringRawBuffer;
}

static LONG g_logs;
static void wh_log(const char *fmt, ...)
{
    char buf[1025];   /* wvsprintf writes at most 1024 characters */
    va_list args;
    if (InterlockedIncrement(&g_logs) > 400) return;
    va_start(args, fmt);
    wvsprintfA(buf, fmt, args);
    va_end(args);
    OutputDebugStringA(buf);
}
static HRESULT notimpl(const char *what) { wh_log("ms365 webhttp: %s not implemented", what); return E_NOTIMPL; }
#define NOTIMPL(fn, what) static HRESULT WINAPI fn(void) { return notimpl(what); }

/* ---- memory and strings ---------------------------------------------------------------------- */
static void *zalloc(SIZE_T n) { return HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, n); }
static void *grow(void *p, SIZE_T n) { return p ? HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, p, n) : zalloc(n); }
static void mfree(void *p) { if (p) HeapFree(GetProcessHeap(), 0, p); }
static WCHAR *wdup_n(const WCHAR *s, SIZE_T n)
{
    WCHAR *d = zalloc((n + 1) * sizeof(WCHAR));
    if (d && n) memcpy(d, s, n * sizeof(WCHAR));
    return d;
}
static WCHAR *wdup(const WCHAR *s) { return s ? wdup_n(s, wcslen(s)) : NULL; }
static WCHAR *wtrim_dup(const WCHAR *s, SIZE_T n)
{
    while (n && (*s == ' ' || *s == '\t')) { s++; n--; }
    while (n && (s[n - 1] == ' ' || s[n - 1] == '\t')) n--;
    return wdup_n(s, n);
}
static WCHAR *from_hstring(HSTRING h)
{
    UINT32 n = 0;
    const WCHAR *p = (h && p_WindowsGetStringRawBuffer) ? p_WindowsGetStringRawBuffer(h, &n) : NULL;
    return wdup_n(p ? p : L"", p ? n : 0);
}
static HRESULT to_hstring(const WCHAR *s, HSTRING *out)
{
    if (!out) return E_POINTER;
    *out = NULL;
    if (!s || !*s) return S_OK;
    return p_WindowsCreateString(s, (UINT32)wcslen(s), out);
}
static WCHAR *wconcat(const WCHAR *a, const WCHAR *sep, const WCHAR *b)
{
    SIZE_T la = wcslen(a), ls = wcslen(sep), lb = wcslen(b);
    WCHAR *d = zalloc((la + ls + lb + 1) * sizeof(WCHAR));
    if (!d) return NULL;
    memcpy(d, a, la * sizeof(WCHAR)); memcpy(d + la, sep, ls * sizeof(WCHAR)); memcpy(d + la + ls, b, lb * sizeof(WCHAR));
    return d;
}
/* HTTP header field names and methods are RFC 7230 tokens; a header value may not contain CR or LF
 * (it would end the header and start another one) */
static BOOL is_token(const WCHAR *s)
{
    if (!s || !*s) return FALSE;
    for (; *s; s++) if (*s <= 32 || *s >= 127 || wcschr(L"\"(),/:;<=>?@[\\]{}", *s)) return FALSE;
    return TRUE;
}
static BOOL is_field_value(const WCHAR *s)
{
    for (; s && *s; s++) if (*s == '\r' || *s == '\n') return FALSE;
    return TRUE;
}
/* UTF-8 bytes of s (no terminator) */
static BYTE *utf8_of(const WCHAR *s, UINT32 *len)
{
    int n = WideCharToMultiByte(CP_UTF8, 0, s, (int)wcslen(s), NULL, 0, NULL, NULL);
    BYTE *d = zalloc(n + 1);
    if (d && n) WideCharToMultiByte(CP_UTF8, 0, s, (int)wcslen(s), (char *)d, n, NULL, NULL);
    *len = d ? (UINT32)n : 0;
    return d;
}
static WCHAR *wide_of_utf8(const BYTE *p, UINT32 len)
{
    if (len >= 3 && p[0] == 0xef && p[1] == 0xbb && p[2] == 0xbf) { p += 3; len -= 3; }
    int n = len ? MultiByteToWideChar(CP_UTF8, 0, (const char *)p, (int)len, NULL, 0) : 0;
    WCHAR *d = zalloc((n + 1) * sizeof(WCHAR));
    if (d && n) MultiByteToWideChar(CP_UTF8, 0, (const char *)p, (int)len, d, n);
    return d;
}

/* call a method through a foreign interface pointer */
#define VCALL(p, idx, type) ((type)((*(void ***)(p))[idx]))
static void unk_release(void *p) { if (p) VCALL(p, 2, ULONG (WINAPI *)(void *))(p); }
static void unk_addref(void *p) { if (p) VCALL(p, 1, ULONG (WINAPI *)(void *))(p); }
static HRESULT unk_qi(void *p, REFIID iid, void **out) { return VCALL(p, 0, HRESULT (WINAPI *)(void *, REFIID, void **))(p, iid, out); }
static void set_unk(IUnknown **slot, void *v) { unk_addref(v); unk_release(*slot); *slot = v; }

/* ---- objects: a set of interfaces sharing one refcount; each interface pointer is a {vtbl, obj}
 * pair so methods can find their object --------------------------------------------------------- */
enum { K_OTHER, K_REQUEST, K_RESPONSE, K_CONTENT, K_HEADERS, K_MAP, K_ISTREAM, K_BUFFER, K_FILTER,
       K_CLIENT, K_METHOD, K_MEDIATYPE, K_READER, K_OP, K_FACTORY };
#define MAXIF 8
struct obj;
struct itf { void *const *vtbl; struct obj *obj; };
struct obj {
    struct itf ifs[MAXIF]; const GUID *iids[MAXIF]; int n;
    LONG ref; BOOL immortal; int kind; const WCHAR *cls;
    void (*destroy)(struct obj *);
};
static void obj_init(struct obj *o, int kind, const WCHAR *cls, void (*destroy)(struct obj *), int n,
                     void *const *const *vtbls, const GUID *const *iids)
{
    o->ref = 1; o->kind = kind; o->cls = cls; o->destroy = destroy; o->n = n;
    for (int i = 0; i < n; i++) { o->ifs[i].vtbl = vtbls[i]; o->ifs[i].obj = o; o->iids[i] = iids[i]; }
}
static ULONG obj_addref(struct obj *o) { return o->immortal ? 2 : (ULONG)InterlockedIncrement(&o->ref); }
static ULONG obj_release(struct obj *o)
{
    if (!o || o->immortal) return 1;
    LONG r = InterlockedDecrement(&o->ref);
    if (!r) { if (o->destroy) o->destroy(o); mfree(o); }
    return (ULONG)r;
}
static void *itf0(struct obj *o) { return &o->ifs[0]; }
static void *ret_itf(struct obj *o, void **out) { *out = o ? itf0(o) : NULL; return *out; }   /* hands over o's reference */

static HRESULT WINAPI o_QueryInterface(struct itf *s, REFIID iid, void **out)
{
    struct obj *o = s->obj;
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &IID_IUnknown) || IsEqualGUID(iid, &IID_Inspectable) || IsEqualGUID(iid, &IID_AgileObject)) {
        *out = itf0(o); obj_addref(o); return S_OK;
    }
    for (int i = 0; i < o->n; i++)
        if (IsEqualGUID(iid, o->iids[i])) { *out = &o->ifs[i]; obj_addref(o); return S_OK; }
    if (IsEqualGUID(iid, &IID_Self)) { *out = o; obj_addref(o); return S_OK; }
    *out = NULL;
    return E_NOINTERFACE;
}
static ULONG WINAPI o_AddRef(struct itf *s) { return obj_addref(s->obj); }
static ULONG WINAPI o_Release(struct itf *s) { return obj_release(s->obj); }
static HRESULT WINAPI o_GetIids(struct itf *s, ULONG *count, GUID **iids) { (void)s; if (count) *count = 0; if (iids) *iids = NULL; return S_OK; }
static HRESULT WINAPI o_GetRuntimeClassName(struct itf *s, HSTRING *name) { return to_hstring(s->obj->cls, name); }
static HRESULT WINAPI o_GetTrustLevel(struct itf *s, int *level) { (void)s; if (level) *level = 0; return S_OK; }
#define INSPECTABLE (void *)o_QueryInterface, (void *)o_AddRef, (void *)o_Release, (void *)o_GetIids, (void *)o_GetRuntimeClassName, (void *)o_GetTrustLevel
#define UNKNOWN     (void *)o_QueryInterface, (void *)o_AddRef, (void *)o_Release

/* our object behind an interface pointer of any instance of this DLL (it is loaded twice, as ole32.dll
 * and as ms365shim.dll), or NULL; the caller keeps its own reference */
static struct obj *ours(void *unk, int kind)
{
    struct obj *o = NULL;
    if (!unk || FAILED(unk_qi(unk, &IID_Self, (void **)&o)) || !o) return NULL;
    obj_release(o);
    return o->kind == kind ? o : NULL;
}

static HRESULT WINAPI w_ok(void) { return S_OK; }
static HRESULT WINAPI w_tostring(struct itf *s, HSTRING *v) { return to_hstring(s->obj->cls, v); }
static void *const closable_vtbl[] = { INSPECTABLE, (void *)w_ok };
static void *const stringable_vtbl[] = { INSPECTABLE, (void *)w_tostring };

/* ---- async operations ----------------------------------------------------------------------- */
enum { ST_STARTED, ST_COMPLETED, ST_CANCELED, ST_ERROR };
enum { R_OBJ, R_STR, R_U32, R_U64 };
struct op {
    struct obj o; SRWLOCK lock;
    INT32 status; HRESULT error; UINT32 id; BOOL fired;
    int rkind; IUnknown *robj; WCHAR *rstr; UINT64 rnum;
    IUnknown *completed, *progress;
    void (*cancel)(void *); void *cancel_ctx;
};
static LONG g_op_ids;
static void op_destroy(struct obj *o)
{
    struct op *op = (struct op *)o;
    unk_release(op->robj); mfree(op->rstr); unk_release(op->completed); unk_release(op->progress);
}
static void op_finish(struct op *op, INT32 status, HRESULT error, IUnknown *robj, const WCHAR *rstr, UINT64 rnum)
{
    IUnknown *handler = NULL;
    AcquireSRWLockExclusive(&op->lock);
    if (op->status != ST_STARTED) { ReleaseSRWLockExclusive(&op->lock); return; }
    op->status = status; op->error = error; op->rnum = rnum;
    if (robj) { unk_addref(robj); op->robj = robj; }
    if (rstr) op->rstr = wdup(rstr);
    if (op->completed && !op->fired) { op->fired = TRUE; handler = op->completed; unk_addref(handler); }
    ReleaseSRWLockExclusive(&op->lock);
    if (handler) {   /* AsyncOperation[WithProgress]CompletedHandler::Invoke(asyncInfo, status) */
        VCALL(handler, 3, HRESULT (WINAPI *)(void *, void *, INT32))(handler, itf0(&op->o), status);
        unk_release(handler);
    }
}
static HRESULT WINAPI op_put_Completed(struct itf *s, IUnknown *handler)
{
    struct op *op = (struct op *)s->obj;
    BOOL fire = FALSE;
    AcquireSRWLockExclusive(&op->lock);
    if (op->completed) { ReleaseSRWLockExclusive(&op->lock); return E_ILLEGAL_DELEGATE_ASSIGNMENT_; }
    unk_addref(handler); op->completed = handler;
    if (handler && op->status != ST_STARTED && !op->fired) fire = op->fired = TRUE;
    INT32 status = op->status;
    ReleaseSRWLockExclusive(&op->lock);
    if (fire) VCALL(handler, 3, HRESULT (WINAPI *)(void *, void *, INT32))(handler, itf0(&op->o), status);
    return S_OK;
}
static HRESULT WINAPI op_get_Completed(struct itf *s, IUnknown **out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    AcquireSRWLockShared(&op->lock); *out = op->completed; unk_addref(*out); ReleaseSRWLockShared(&op->lock);
    return S_OK;
}
static HRESULT WINAPI op_put_Progress(struct itf *s, IUnknown *handler)
{
    struct op *op = (struct op *)s->obj;
    AcquireSRWLockExclusive(&op->lock); set_unk(&op->progress, handler); ReleaseSRWLockExclusive(&op->lock);
    return S_OK;
}
static HRESULT WINAPI op_get_Progress(struct itf *s, IUnknown **out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    AcquireSRWLockShared(&op->lock); *out = op->progress; unk_addref(*out); ReleaseSRWLockShared(&op->lock);
    return S_OK;
}
static HRESULT op_results_check(struct op *op)
{
    switch (op->status) {
    case ST_COMPLETED: return S_OK;
    case ST_ERROR: return op->error;
    case ST_CANCELED: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    default: return E_ILLEGAL_METHOD_CALL_;
    }
}
static HRESULT WINAPI op_results_obj(struct itf *s, IUnknown **out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    *out = NULL;
    AcquireSRWLockShared(&op->lock);
    HRESULT hr = op_results_check(op);
    if (hr == S_OK) { *out = op->robj; unk_addref(*out); }
    ReleaseSRWLockShared(&op->lock);
    return hr;
}
static HRESULT WINAPI op_results_str(struct itf *s, HSTRING *out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    *out = NULL;
    AcquireSRWLockShared(&op->lock);
    HRESULT hr = op_results_check(op);
    if (hr == S_OK) hr = to_hstring(op->rstr, out);
    ReleaseSRWLockShared(&op->lock);
    return hr;
}
static HRESULT WINAPI op_results_u32(struct itf *s, UINT32 *out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    AcquireSRWLockShared(&op->lock);
    HRESULT hr = op_results_check(op);
    *out = hr == S_OK ? (UINT32)op->rnum : 0;
    ReleaseSRWLockShared(&op->lock);
    return hr;
}
static HRESULT WINAPI op_results_u64(struct itf *s, UINT64 *out)
{
    struct op *op = (struct op *)s->obj;
    if (!out) return E_POINTER;
    AcquireSRWLockShared(&op->lock);
    HRESULT hr = op_results_check(op);
    *out = hr == S_OK ? op->rnum : 0;
    ReleaseSRWLockShared(&op->lock);
    return hr;
}
static HRESULT WINAPI ai_get_Id(struct itf *s, UINT32 *v) { if (!v) return E_POINTER; *v = ((struct op *)s->obj)->id; return S_OK; }
static HRESULT WINAPI ai_get_Status(struct itf *s, INT32 *v)
{
    struct op *op = (struct op *)s->obj;
    if (!v) return E_POINTER;
    AcquireSRWLockShared(&op->lock); *v = op->status; ReleaseSRWLockShared(&op->lock);
    return S_OK;
}
static HRESULT WINAPI ai_get_ErrorCode(struct itf *s, HRESULT *v)
{
    struct op *op = (struct op *)s->obj;
    if (!v) return E_POINTER;
    AcquireSRWLockShared(&op->lock);
    *v = op->status == ST_ERROR ? op->error : op->status == ST_CANCELED ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    ReleaseSRWLockShared(&op->lock);
    return S_OK;
}
static HRESULT WINAPI ai_Cancel(struct itf *s)
{
    struct op *op = (struct op *)s->obj;
    AcquireSRWLockExclusive(&op->lock);
    if (op->status == ST_STARTED && op->cancel) op->cancel(op->cancel_ctx);
    ReleaseSRWLockExclusive(&op->lock);
    return S_OK;
}
static void *const asyncinfo_vtbl[] = { INSPECTABLE, (void *)ai_get_Id, (void *)ai_get_Status, (void *)ai_get_ErrorCode, (void *)ai_Cancel, (void *)w_ok };
/* IAsyncOperationWithProgress<T, P>: Progress, Completed, GetResults; IAsyncOperation<T>: Completed, GetResults */
static void *const op_obj_vtbl[] = { INSPECTABLE, (void *)op_put_Progress, (void *)op_get_Progress, (void *)op_put_Completed, (void *)op_get_Completed, (void *)op_results_obj };
static void *const op_str_vtbl[] = { INSPECTABLE, (void *)op_put_Progress, (void *)op_get_Progress, (void *)op_put_Completed, (void *)op_get_Completed, (void *)op_results_str };
static void *const op_u64_vtbl[] = { INSPECTABLE, (void *)op_put_Progress, (void *)op_get_Progress, (void *)op_put_Completed, (void *)op_get_Completed, (void *)op_results_u64 };
static void *const op_u32_vtbl[] = { INSPECTABLE, (void *)op_put_Completed, (void *)op_get_Completed, (void *)op_results_u32 };

static struct op *op_new(const GUID *iid, int rkind)
{
    struct op *op = zalloc(sizeof(*op));
    if (!op) return NULL;
    void *const *vt[] = { rkind == R_U32 ? op_u32_vtbl : rkind == R_U64 ? op_u64_vtbl : rkind == R_STR ? op_str_vtbl : op_obj_vtbl, asyncinfo_vtbl };
    const GUID *ii[] = { iid, &IID_AsyncInfo };
    obj_init(&op->o, K_OP, L"Windows.Foundation.IAsyncOperation", op_destroy, 2, vt, ii);
    InitializeSRWLock(&op->lock);
    op->rkind = rkind; op->id = (UINT32)InterlockedIncrement(&g_op_ids);
    return op;
}
/* an operation that has already completed with its result */
static HRESULT op_done(const GUID *iid, int rkind, IUnknown *robj, const WCHAR *rstr, UINT64 rnum, void **out)
{
    struct op *op = op_new(iid, rkind);
    if (!out) { obj_release(op ? &op->o : NULL); return E_POINTER; }
    *out = NULL;
    if (!op) return E_OUTOFMEMORY;
    op_finish(op, ST_COMPLETED, S_OK, robj, rstr, rnum);
    ret_itf(&op->o, out);
    return S_OK;
}

/* ---- key/value lists: header collections (case-insensitive, string values) and property maps
 * (case-sensitive, object values) ---------------------------------------------------------------- */
struct kv { WCHAR *key; WCHAR *str; IUnknown *unk; };
struct kvlist { struct kv *v; UINT32 n, cap; BOOL icase; };
static void kv_clear(struct kvlist *l)
{
    for (UINT32 i = 0; i < l->n; i++) { mfree(l->v[i].key); mfree(l->v[i].str); unk_release(l->v[i].unk); }
    mfree(l->v); l->v = NULL; l->n = l->cap = 0;
}
static int kv_find(const struct kvlist *l, const WCHAR *key)
{
    for (UINT32 i = 0; i < l->n; i++)
        if ((l->icase ? _wcsicmp(l->v[i].key, key) : wcscmp(l->v[i].key, key)) == 0) return (int)i;
    return -1;
}
/* replace (or, with append, join a header's values with ", "); returns whether the key existed */
static BOOL kv_set(struct kvlist *l, const WCHAR *key, const WCHAR *str, IUnknown *unk, BOOL append)
{
    int i = kv_find(l, key);
    if (i >= 0) {
        struct kv *e = &l->v[i];
        WCHAR *v = (append && e->str && e->str[0] && str) ? wconcat(e->str, L", ", str) : wdup(str);
        mfree(e->str); e->str = v;
        set_unk(&e->unk, unk);
        return TRUE;
    }
    if (l->n == l->cap) {
        UINT32 cap = l->cap ? l->cap * 2 : 8;
        struct kv *v = grow(l->v, cap * sizeof(*v));
        if (!v) return FALSE;
        l->v = v; l->cap = cap;
    }
    struct kv *e = &l->v[l->n++];
    e->key = wdup(key); e->str = wdup(str); e->unk = unk; unk_addref(unk);
    return FALSE;
}
static BOOL kv_remove(struct kvlist *l, const WCHAR *key)
{
    int i = kv_find(l, key);
    if (i < 0) return FALSE;
    mfree(l->v[i].key); mfree(l->v[i].str); unk_release(l->v[i].unk);
    memmove(&l->v[i], &l->v[i + 1], (l->n - i - 1) * sizeof(l->v[0]));
    l->n--;
    return TRUE;
}
static void kv_copy(struct kvlist *dst, const struct kvlist *src)
{
    dst->icase = src->icase;
    for (UINT32 i = 0; i < src->n; i++) kv_set(dst, src->v[i].key, src->v[i].str, src->v[i].unk, FALSE);
}

/* a pair handed out by iterators */
struct kvp { struct obj o; WCHAR *key; WCHAR *str; IUnknown *unk; };
static void kvp_destroy(struct obj *o) { struct kvp *p = (struct kvp *)o; mfree(p->key); mfree(p->str); unk_release(p->unk); }
static HRESULT WINAPI kvp_get_Key(struct itf *s, HSTRING *v) { return to_hstring(((struct kvp *)s->obj)->key, v); }
static HRESULT WINAPI kvp_get_StrValue(struct itf *s, HSTRING *v) { return to_hstring(((struct kvp *)s->obj)->str, v); }
static HRESULT WINAPI kvp_get_ObjValue(struct itf *s, IUnknown **v)
{
    if (!v) return E_POINTER;
    *v = ((struct kvp *)s->obj)->unk; unk_addref(*v);
    return S_OK;
}
static void *const kvp_ss_vtbl[] = { INSPECTABLE, (void *)kvp_get_Key, (void *)kvp_get_StrValue };
static void *const kvp_so_vtbl[] = { INSPECTABLE, (void *)kvp_get_Key, (void *)kvp_get_ObjValue };
static struct obj *kvp_new(const struct kv *e, BOOL objects)
{
    struct kvp *p = zalloc(sizeof(*p));
    if (!p) return NULL;
    void *const *vt[] = { objects ? kvp_so_vtbl : kvp_ss_vtbl };
    const GUID *ii[] = { objects ? &IID_KvpSO : &IID_KvpSS };
    obj_init(&p->o, K_OTHER, L"Windows.Foundation.Collections.IKeyValuePair", kvp_destroy, 1, vt, ii);
    p->key = wdup(e->key); p->str = wdup(e->str); p->unk = e->unk; unk_addref(p->unk);
    return &p->o;
}

/* iterator over a snapshot */
struct iter { struct obj o; struct kvlist l; UINT32 pos; BOOL objects; };
static void iter_destroy(struct obj *o) { kv_clear(&((struct iter *)o)->l); }
static HRESULT WINAPI it_get_Current(struct itf *s, void **out)
{
    struct iter *it = (struct iter *)s->obj;
    if (!out) return E_POINTER;
    *out = NULL;
    if (it->pos >= it->l.n) return E_BOUNDS_;
    struct obj *p = kvp_new(&it->l.v[it->pos], it->objects);
    if (!p) return E_OUTOFMEMORY;
    ret_itf(p, out);
    return S_OK;
}
static HRESULT WINAPI it_get_HasCurrent(struct itf *s, boolean8 *v)
{
    struct iter *it = (struct iter *)s->obj;
    if (!v) return E_POINTER;
    *v = it->pos < it->l.n;
    return S_OK;
}
static HRESULT WINAPI it_MoveNext(struct itf *s, boolean8 *v)
{
    struct iter *it = (struct iter *)s->obj;
    if (it->pos < it->l.n) it->pos++;
    if (v) *v = it->pos < it->l.n;
    return S_OK;
}
static HRESULT WINAPI it_GetMany(struct itf *s, UINT32 capacity, void **items, UINT32 *actual)
{
    struct iter *it = (struct iter *)s->obj;
    UINT32 n = 0;
    if (!actual || (capacity && !items)) return E_POINTER;
    while (n < capacity && it->pos < it->l.n) {
        struct obj *p = kvp_new(&it->l.v[it->pos], it->objects);
        if (!p) break;
        items[n++] = itf0(p); it->pos++;
    }
    *actual = n;
    return S_OK;
}
static void *const iter_vtbl[] = { INSPECTABLE, (void *)it_get_Current, (void *)it_get_HasCurrent, (void *)it_MoveNext, (void *)it_GetMany };
static HRESULT iter_new(const struct kvlist *l, BOOL objects, void **out)
{
    struct iter *it = zalloc(sizeof(*it));
    if (!out) { mfree(it); return E_POINTER; }
    *out = NULL;
    if (!it) return E_OUTOFMEMORY;
    void *const *vt[] = { iter_vtbl };
    const GUID *ii[] = { objects ? &IID_IteratorSO : &IID_IteratorSS };
    obj_init(&it->o, K_OTHER, L"Windows.Foundation.Collections.IIterator", iter_destroy, 1, vt, ii);
    kv_copy(&it->l, l); it->objects = objects;
    ret_itf(&it->o, out);
    return S_OK;
}

/* maps: header collections and HttpRequestMessage.Properties share this part; GetView hands out a
 * snapshot with the same layout */
struct map { struct obj o; SRWLOCK lock; struct kvlist l; BOOL objects; };
static void map_destroy(struct obj *o) { kv_clear(&((struct map *)o)->l); }
static HRESULT WINAPI map_Lookup(struct itf *s, HSTRING key, void **out)
{
    struct map *m = (struct map *)s->obj;
    WCHAR *k = from_hstring(key);
    HRESULT hr = E_BOUNDS_;
    if (!out) { mfree(k); return E_POINTER; }
    *out = NULL;
    AcquireSRWLockShared(&m->lock);
    int i = k ? kv_find(&m->l, k) : -1;
    if (i >= 0) {
        if (m->objects) { *out = m->l.v[i].unk; unk_addref(*out); hr = S_OK; }
        else hr = to_hstring(m->l.v[i].str, (HSTRING *)out);
    }
    ReleaseSRWLockShared(&m->lock);
    mfree(k);
    return hr;
}
static HRESULT WINAPI map_get_Size(struct itf *s, UINT32 *v)
{
    struct map *m = (struct map *)s->obj;
    if (!v) return E_POINTER;
    AcquireSRWLockShared(&m->lock); *v = m->l.n; ReleaseSRWLockShared(&m->lock);
    return S_OK;
}
static HRESULT WINAPI map_HasKey(struct itf *s, HSTRING key, boolean8 *v)
{
    struct map *m = (struct map *)s->obj;
    WCHAR *k = from_hstring(key);
    if (!v) { mfree(k); return E_POINTER; }
    AcquireSRWLockShared(&m->lock); *v = k && kv_find(&m->l, k) >= 0; ReleaseSRWLockShared(&m->lock);
    mfree(k);
    return S_OK;
}
static HRESULT WINAPI map_Insert(struct itf *s, HSTRING key, void *value, boolean8 *replaced)
{
    struct map *m = (struct map *)s->obj;
    WCHAR *k = from_hstring(key), *v = m->objects ? NULL : from_hstring((HSTRING)value);
    if (!k || (!m->objects && !v)) { mfree(k); mfree(v); return E_OUTOFMEMORY; }
    if (!m->objects && (!is_token(k) || !is_field_value(v))) { mfree(k); mfree(v); return E_INVALIDARG; }
    AcquireSRWLockExclusive(&m->lock);
    BOOL existed = kv_set(&m->l, k, v, m->objects ? value : NULL, FALSE);
    ReleaseSRWLockExclusive(&m->lock);
    if (replaced) *replaced = existed;
    mfree(k); mfree(v);
    return S_OK;
}
static HRESULT WINAPI map_Remove(struct itf *s, HSTRING key)
{
    struct map *m = (struct map *)s->obj;
    WCHAR *k = from_hstring(key);
    AcquireSRWLockExclusive(&m->lock); if (k) kv_remove(&m->l, k); ReleaseSRWLockExclusive(&m->lock);
    mfree(k);
    return S_OK;
}
static HRESULT WINAPI map_Clear(struct itf *s)
{
    struct map *m = (struct map *)s->obj;
    AcquireSRWLockExclusive(&m->lock); kv_clear(&m->l); ReleaseSRWLockExclusive(&m->lock);
    return S_OK;
}
static HRESULT WINAPI map_First(struct itf *s, void **out)
{
    struct map *m = (struct map *)s->obj;
    AcquireSRWLockShared(&m->lock);
    HRESULT hr = iter_new(&m->l, m->objects, out);
    ReleaseSRWLockShared(&m->lock);
    return hr;
}
static HRESULT WINAPI view_Split(struct itf *s, void **first, void **second)
{
    (void)s;
    if (first) *first = NULL;
    if (second) *second = NULL;
    return S_OK;
}
static void *const view_vtbl[] = { INSPECTABLE, (void *)map_Lookup, (void *)map_get_Size, (void *)map_HasKey, (void *)view_Split };
static void *const iterable_vtbl[] = { INSPECTABLE, (void *)map_First };
static HRESULT WINAPI map_GetView(struct itf *s, void **out)
{
    struct map *m = (struct map *)s->obj, *v = zalloc(sizeof(*v));
    if (!out) { mfree(v); return E_POINTER; }
    *out = NULL;
    if (!v) return E_OUTOFMEMORY;
    void *const *vt[] = { view_vtbl, iterable_vtbl };
    const GUID *ii[] = { m->objects ? &IID_MapViewSO : &IID_MapViewSS, m->objects ? &IID_IterableSO : &IID_IterableSS };
    obj_init(&v->o, K_OTHER, L"Windows.Foundation.Collections.IMapView", map_destroy, 2, vt, ii);
    InitializeSRWLock(&v->lock); v->objects = m->objects;
    AcquireSRWLockShared(&m->lock); kv_copy(&v->l, &m->l); ReleaseSRWLockShared(&m->lock);
    ret_itf(&v->o, out);
    return S_OK;
}
static void *const map_vtbl[] = { INSPECTABLE, (void *)map_Lookup, (void *)map_get_Size, (void *)map_HasKey, (void *)map_GetView,
                                  (void *)map_Insert, (void *)map_Remove, (void *)map_Clear };

/* HttpRequestMessage.Properties: IMap<String, Object> */
static struct map *propmap_new(void)
{
    struct map *m = zalloc(sizeof(*m));
    if (!m) return NULL;
    void *const *vt[] = { map_vtbl, iterable_vtbl };
    const GUID *ii[] = { &IID_MapSO, &IID_IterableSO };
    obj_init(&m->o, K_MAP, L"Windows.Foundation.Collections.PropertySet", map_destroy, 2, vt, ii);
    InitializeSRWLock(&m->lock); m->objects = TRUE;
    return m;
}

/* ---- header collections ------------------------------------------------------------------------ */
enum { H_REQUEST, H_RESPONSE, H_CONTENT };
struct headers { struct map m; int which; WCHAR *base; /* response: the request URI, for a relative Location */ };
static void headers_destroy(struct obj *o) { map_destroy(o); mfree(((struct headers *)o)->base); }
static WCHAR *hdr_get(struct headers *h, const WCHAR *name)
{
    AcquireSRWLockShared(&h->m.lock);
    int i = kv_find(&h->m.l, name);
    WCHAR *v = i >= 0 ? wdup(h->m.l.v[i].str) : NULL;
    ReleaseSRWLockShared(&h->m.lock);
    return v;
}
static HRESULT hdr_set(struct headers *h, const WCHAR *name, const WCHAR *value, BOOL append)
{
    if (!is_token(name) || !is_field_value(value)) return E_INVALIDARG;
    AcquireSRWLockExclusive(&h->m.lock);
    if (value) kv_set(&h->m.l, name, value, NULL, append);
    else kv_remove(&h->m.l, name);
    ReleaseSRWLockExclusive(&h->m.lock);
    return S_OK;
}
static HRESULT WINAPI hdr_Append(struct itf *s, HSTRING name, HSTRING value)
{
    WCHAR *n = from_hstring(name), *v = from_hstring(value);
    HRESULT hr = (n && v) ? hdr_set((struct headers *)s->obj, n, v, TRUE) : E_OUTOFMEMORY;
    mfree(n); mfree(v);
    return hr;
}
static HRESULT WINAPI hdr_TryAppendWithoutValidation(struct itf *s, HSTRING name, HSTRING value, boolean8 *ok)
{
    HRESULT hr = hdr_Append(s, name, value);
    if (ok) *ok = SUCCEEDED(hr);
    return S_OK;
}
static HRESULT WINAPI hdr_ToString(struct itf *s, HSTRING *out)
{
    struct headers *h = (struct headers *)s->obj;
    WCHAR *text = wdup(L"");
    AcquireSRWLockShared(&h->m.lock);
    for (UINT32 i = 0; i < h->m.l.n && text; i++) {
        WCHAR *line = wconcat(h->m.l.v[i].key, L": ", h->m.l.v[i].str ? h->m.l.v[i].str : L"");
        WCHAR *t = line ? wconcat(text, line, L"\r\n") : NULL;
        mfree(line); mfree(text); text = t;
    }
    ReleaseSRWLockShared(&h->m.lock);
    HRESULT hr = text ? to_hstring(text, out) : E_OUTOFMEMORY;
    mfree(text);
    return hr;
}
static void *const hdr_stringable_vtbl[] = { INSPECTABLE, (void *)hdr_ToString };

/* a Uri (Wine's Windows.Foundation.Uri) for an absolute URI string */
static HRESULT make_uri(const WCHAR *s, void **out)
{
    HSTRING cls = NULL, str = NULL;
    void *factory = NULL;
    HRESULT hr;
    *out = NULL;
    if (!p_RoGetActivationFactory) return E_NOTIMPL;
    if (FAILED(hr = p_WindowsCreateString(L"Windows.Foundation.Uri", 22, &cls))) return hr;
    hr = p_RoGetActivationFactory(cls, &IID_UriFactory, &factory);
    if (SUCCEEDED(hr) && SUCCEEDED(hr = to_hstring(s, &str)))
        hr = VCALL(factory, 6, HRESULT (WINAPI *)(void *, HSTRING, void **))(factory, str, out);   /* CreateUri */
    if (factory) unk_release(factory);
    if (str) p_WindowsDeleteString(str);
    p_WindowsDeleteString(cls);
    return hr;
}
static WCHAR *uri_string(void *uri)   /* IUriRuntimeClass::get_AbsoluteUri */
{
    HSTRING h = NULL;
    if (!uri || FAILED(VCALL(uri, 6, HRESULT (WINAPI *)(void *, HSTRING *))(uri, &h))) return NULL;
    WCHAR *s = from_hstring(h);
    if (h) p_WindowsDeleteString(h);
    return s;
}
/* resolve a Location header against the request URI */
static WCHAR *resolve_location(const WCHAR *base, const WCHAR *loc)
{
    const WCHAR *p = wcsstr(base ? base : L"", L"://");
    if (wcsstr(loc, L"://") || !base || !p) return wdup(loc);
    SIZE_T scheme = p - base, auth = scheme + 3;
    while (base[auth] && base[auth] != '/' && base[auth] != '?' && base[auth] != '#') auth++;
    if (loc[0] == '/' && loc[1] == '/') { WCHAR *s = wdup_n(base, scheme + 1); WCHAR *r = s ? wconcat(s, L"", loc) : NULL; mfree(s); return r; }
    if (loc[0] == '/') { WCHAR *s = wdup_n(base, auth); WCHAR *r = s ? wconcat(s, L"", loc) : NULL; mfree(s); return r; }
    SIZE_T end = auth, slash = auth;
    while (base[end] && base[end] != '?' && base[end] != '#') { if (base[end] == '/') slash = end + 1; end++; }
    if (loc[0] == '?') { WCHAR *s = wdup_n(base, end); WCHAR *r = s ? wconcat(s, slash == auth ? L"/" : L"", loc) : NULL; mfree(s); return r; }
    if (slash == auth) { WCHAR *s = wdup_n(base, auth); WCHAR *r = s ? wconcat(s, L"/", loc) : NULL; mfree(s); return r; }
    WCHAR *s = wdup_n(base, slash);
    WCHAR *r = s ? wconcat(s, L"", loc) : NULL;
    mfree(s);
    return r;
}
static HRESULT WINAPI resp_hdr_get_Location(struct itf *s, void **out)
{
    struct headers *h = (struct headers *)s->obj;
    if (!out) return E_POINTER;
    *out = NULL;
    WCHAR *loc = hdr_get(h, L"Location");
    if (!loc) return S_OK;
    WCHAR *abs = resolve_location(h->base, loc);
    HRESULT hr = abs ? make_uri(abs, out) : E_OUTOFMEMORY;
    mfree(loc); mfree(abs);
    return hr;
}
static HRESULT WINAPI resp_hdr_put_Location(struct itf *s, void *uri)
{
    WCHAR *v = uri_string(uri);
    HRESULT hr = hdr_set((struct headers *)s->obj, L"Location", v, FALSE);
    mfree(v);
    return hr;
}

/* IReference<UInt64> for HttpContentHeaderCollection.ContentLength */
struct u64ref { struct obj o; UINT64 v; };
static HRESULT WINAPI u64ref_get_Value(struct itf *s, UINT64 *v) { if (!v) return E_POINTER; *v = ((struct u64ref *)s->obj)->v; return S_OK; }
static void *const u64ref_vtbl[] = { INSPECTABLE, (void *)u64ref_get_Value };

/* HttpMediaTypeHeaderValue: "media/type" plus the parameters as written */
struct mediatype { struct obj o; SRWLOCK lock; WCHAR *media; WCHAR *params; };
static void mediatype_destroy(struct obj *o) { mfree(((struct mediatype *)o)->media); mfree(((struct mediatype *)o)->params); }
static WCHAR *param_value(const WCHAR *params, const WCHAR *name)
{
    SIZE_T nl = wcslen(name);
    for (const WCHAR *p = params; p && *p; ) {
        while (*p == ' ' || *p == ';') p++;
        const WCHAR *e = p;
        while (*e && *e != ';') e++;
        if ((SIZE_T)(e - p) > nl && !_wcsnicmp(p, name, nl) && p[nl] == '=') {
            const WCHAR *v = p + nl + 1;
            SIZE_T n = e - v;
            while (n && v[n - 1] == ' ') n--;
            if (n >= 2 && v[0] == '"' && v[n - 1] == '"') { v++; n -= 2; }
            return wdup_n(v, n);
        }
        p = e;
    }
    return NULL;
}
static WCHAR *params_without(const WCHAR *params, const WCHAR *name)
{
    SIZE_T nl = wcslen(name);
    WCHAR *out = wdup(L"");
    for (const WCHAR *p = params; p && *p && out; ) {
        while (*p == ' ' || *p == ';') p++;
        const WCHAR *e = p;
        while (*e && *e != ';') e++;
        if (e > p && !((SIZE_T)(e - p) > nl && !_wcsnicmp(p, name, nl) && p[nl] == '=')) {
            WCHAR *part = wtrim_dup(p, e - p);
            WCHAR *t = part ? wconcat(out, out[0] ? L"; " : L"", part) : NULL;
            mfree(part); mfree(out); out = t;
        }
        p = e;
    }
    return out;
}
static WCHAR *mediatype_text(struct mediatype *m)
{
    AcquireSRWLockShared(&m->lock);
    WCHAR *t = (m->params && m->params[0]) ? wconcat(m->media, L"; ", m->params) : wdup(m->media);
    ReleaseSRWLockShared(&m->lock);
    return t;
}
static HRESULT WINAPI mt_get_CharSet(struct itf *s, HSTRING *out)
{
    struct mediatype *m = (struct mediatype *)s->obj;
    AcquireSRWLockShared(&m->lock); WCHAR *v = param_value(m->params, L"charset"); ReleaseSRWLockShared(&m->lock);
    HRESULT hr = to_hstring(v, out);
    mfree(v);
    return hr;
}
static HRESULT WINAPI mt_put_CharSet(struct itf *s, HSTRING value)
{
    struct mediatype *m = (struct mediatype *)s->obj;
    WCHAR *v = from_hstring(value);
    AcquireSRWLockExclusive(&m->lock);
    WCHAR *rest = params_without(m->params, L"charset"), *p = NULL;
    if (rest && v && v[0]) { WCHAR *cs = wconcat(L"charset", L"=", v); p = cs ? wconcat(cs, rest[0] ? L"; " : L"", rest) : NULL; mfree(cs); mfree(rest); }
    else p = rest;
    mfree(m->params); m->params = p;
    ReleaseSRWLockExclusive(&m->lock);
    mfree(v);
    return S_OK;
}
static HRESULT WINAPI mt_get_MediaType(struct itf *s, HSTRING *out)
{
    struct mediatype *m = (struct mediatype *)s->obj;
    AcquireSRWLockShared(&m->lock); HRESULT hr = to_hstring(m->media, out); ReleaseSRWLockShared(&m->lock);
    return hr;
}
static HRESULT WINAPI mt_put_MediaType(struct itf *s, HSTRING value)
{
    struct mediatype *m = (struct mediatype *)s->obj;
    WCHAR *v = from_hstring(value);
    if (!v || !wcschr(v, '/')) { mfree(v); return E_INVALIDARG; }
    AcquireSRWLockExclusive(&m->lock); mfree(m->media); m->media = v; ReleaseSRWLockExclusive(&m->lock);
    return S_OK;
}
NOTIMPL(mt_get_Parameters, "HttpMediaTypeHeaderValue.Parameters")
static HRESULT WINAPI mt_ToString(struct itf *s, HSTRING *out)
{
    WCHAR *t = mediatype_text((struct mediatype *)s->obj);
    HRESULT hr = t ? to_hstring(t, out) : E_OUTOFMEMORY;
    mfree(t);
    return hr;
}
static void *const mediatype_vtbl[] = { INSPECTABLE, (void *)mt_get_CharSet, (void *)mt_put_CharSet, (void *)mt_get_MediaType, (void *)mt_put_MediaType, (void *)mt_get_Parameters };
static void *const mediatype_stringable_vtbl[] = { INSPECTABLE, (void *)mt_ToString };
/* "type/subtype; param=value; ..." */
static HRESULT mediatype_parse(const WCHAR *text, void **out)
{
    *out = NULL;
    const WCHAR *semi = wcschr(text, ';');
    WCHAR *media = wtrim_dup(text, semi ? (SIZE_T)(semi - text) : wcslen(text));
    if (!media || !media[0] || !wcschr(media, '/') || wcschr(media, ' ')) { mfree(media); return E_INVALIDARG; }
    struct mediatype *m = zalloc(sizeof(*m));
    if (!m) { mfree(media); return E_OUTOFMEMORY; }
    void *const *vt[] = { mediatype_vtbl, mediatype_stringable_vtbl };
    const GUID *ii[] = { &IID_MediaType, &IID_Stringable };
    obj_init(&m->o, K_MEDIATYPE, L"Windows.Web.Http.Headers.HttpMediaTypeHeaderValue", mediatype_destroy, 2, vt, ii);
    InitializeSRWLock(&m->lock);
    m->media = media; m->params = semi ? params_without(semi + 1, L"") : NULL;
    ret_itf(&m->o, out);
    return S_OK;
}

/* HttpContentCodingHeaderValueCollection: just enough to add a Content-Encoding */
struct codings { struct obj o; struct headers *h; };
static void codings_destroy(struct obj *o) { obj_release(&((struct codings *)o)->h->m.o); }
static HRESULT WINAPI cod_ParseAdd(struct itf *s, HSTRING input)
{
    WCHAR *v = from_hstring(input);
    if (!v || !v[0]) { mfree(v); return E_INVALIDARG; }
    HRESULT hr = hdr_set(((struct codings *)s->obj)->h, L"Content-Encoding", v, TRUE);
    mfree(v);
    return hr;
}
static HRESULT WINAPI cod_TryParseAdd(struct itf *s, HSTRING input, boolean8 *ok)
{
    HRESULT hr = cod_ParseAdd(s, input);
    if (ok) *ok = SUCCEEDED(hr);
    return S_OK;
}
static HRESULT WINAPI cod_ToString(struct itf *s, HSTRING *out)
{
    WCHAR *v = hdr_get(((struct codings *)s->obj)->h, L"Content-Encoding");
    HRESULT hr = to_hstring(v, out);
    mfree(v);
    return hr;
}
static void *const codings_vtbl[] = { INSPECTABLE, (void *)cod_ParseAdd, (void *)cod_TryParseAdd };
static void *const codings_stringable_vtbl[] = { INSPECTABLE, (void *)cod_ToString };

static HRESULT WINAPI cnt_hdr_get_ContentEncoding(struct itf *s, void **out)
{
    struct codings *c = zalloc(sizeof(*c));
    if (!out) { mfree(c); return E_POINTER; }
    *out = NULL;
    if (!c) return E_OUTOFMEMORY;
    void *const *vt[] = { codings_vtbl, codings_stringable_vtbl };
    const GUID *ii[] = { &IID_ContentCodings, &IID_Stringable };
    obj_init(&c->o, K_OTHER, L"Windows.Web.Http.Headers.HttpContentCodingHeaderValueCollection", codings_destroy, 2, vt, ii);
    c->h = (struct headers *)s->obj; obj_addref(&c->h->m.o);
    ret_itf(&c->o, out);
    return S_OK;
}
static HRESULT WINAPI cnt_hdr_get_ContentLength(struct itf *s, void **out)
{
    if (!out) return E_POINTER;
    *out = NULL;
    WCHAR *v = hdr_get((struct headers *)s->obj, L"Content-Length");
    if (!v) return S_OK;
    struct u64ref *r = zalloc(sizeof(*r));
    if (!r) { mfree(v); return E_OUTOFMEMORY; }
    void *const *vt[] = { u64ref_vtbl };
    const GUID *ii[] = { &IID_RefU64 };
    obj_init(&r->o, K_OTHER, L"Windows.Foundation.IReference`1<UInt64>", NULL, 1, vt, ii);
    r->v = _wcstoui64(v, NULL, 10);
    mfree(v);
    ret_itf(&r->o, out);
    return S_OK;
}
static HRESULT WINAPI cnt_hdr_put_ContentLength(struct itf *s, void *ref)
{
    UINT64 n = 0;
    WCHAR buf[24];
    if (ref && FAILED(VCALL(ref, 6, HRESULT (WINAPI *)(void *, UINT64 *))(ref, &n))) return E_INVALIDARG;   /* IReference::get_Value */
    if (ref) _ui64tow(n, buf, 10);
    return hdr_set((struct headers *)s->obj, L"Content-Length", ref ? buf : NULL, FALSE);
}
static HRESULT WINAPI cnt_hdr_get_ContentType(struct itf *s, void **out)
{
    if (!out) return E_POINTER;
    *out = NULL;
    WCHAR *v = hdr_get((struct headers *)s->obj, L"Content-Type");
    HRESULT hr = v ? mediatype_parse(v, out) : S_OK;
    mfree(v);
    return hr == E_INVALIDARG ? S_OK : hr;
}
static HRESULT WINAPI cnt_hdr_put_ContentType(struct itf *s, void *value)
{
    HSTRING str = NULL;
    void *sb = NULL;
    WCHAR *v = NULL;
    if (value && SUCCEEDED(unk_qi(value, &IID_Stringable, &sb))) {
        if (SUCCEEDED(VCALL(sb, 6, HRESULT (WINAPI *)(void *, HSTRING *))(sb, &str))) v = from_hstring(str);
        if (str) p_WindowsDeleteString(str);
        unk_release(sb);
    }
    if (value && !v) return E_INVALIDARG;
    HRESULT hr = hdr_set((struct headers *)s->obj, L"Content-Type", v, FALSE);
    mfree(v);
    return hr;
}
NOTIMPL(req_hdr_typed, "HttpRequestHeaderCollection typed header property")
NOTIMPL(resp_hdr_typed, "HttpResponseHeaderCollection typed header property")
NOTIMPL(cnt_hdr_typed, "HttpContentHeaderCollection typed header property")
static void *const req_hdr_vtbl[] = { INSPECTABLE,
    (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed,   /* Accept .. put_Authorization */
    (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed,   /* CacheControl .. put_Date */
    (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed,   /* Expect .. put_Host */
    (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed,   /* IfModifiedSince .. get_MaxForwards */
    (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed, (void *)req_hdr_typed,   /* put_MaxForwards .. put_Referer */
    (void *)req_hdr_typed, (void *)req_hdr_typed,                                                                       /* TransferEncoding, UserAgent */
    (void *)hdr_Append, (void *)hdr_TryAppendWithoutValidation };
static void *const resp_hdr_vtbl[] = { INSPECTABLE,
    (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed,   /* Age .. Connection */
    (void *)resp_hdr_typed, (void *)resp_hdr_typed,                                                                          /* Date */
    (void *)resp_hdr_get_Location, (void *)resp_hdr_put_Location,
    (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed, (void *)resp_hdr_typed,   /* ProxyAuthenticate .. WwwAuthenticate */
    (void *)hdr_Append, (void *)hdr_TryAppendWithoutValidation };
static void *const cnt_hdr_vtbl[] = { INSPECTABLE,
    (void *)cnt_hdr_typed, (void *)cnt_hdr_typed,                          /* ContentDisposition */
    (void *)cnt_hdr_get_ContentEncoding, (void *)cnt_hdr_typed,             /* ContentEncoding, ContentLanguage */
    (void *)cnt_hdr_get_ContentLength, (void *)cnt_hdr_put_ContentLength,
    (void *)cnt_hdr_typed, (void *)cnt_hdr_typed, (void *)cnt_hdr_typed, (void *)cnt_hdr_typed,   /* ContentLocation, ContentMD5 */
    (void *)cnt_hdr_typed, (void *)cnt_hdr_typed,                          /* ContentRange */
    (void *)cnt_hdr_get_ContentType, (void *)cnt_hdr_put_ContentType,
    (void *)cnt_hdr_typed, (void *)cnt_hdr_typed, (void *)cnt_hdr_typed, (void *)cnt_hdr_typed,   /* Expires, LastModified */
    (void *)hdr_Append, (void *)hdr_TryAppendWithoutValidation };
static struct headers *headers_new(int which)
{
    static const WCHAR *const names[] = { L"Windows.Web.Http.Headers.HttpRequestHeaderCollection",
        L"Windows.Web.Http.Headers.HttpResponseHeaderCollection", L"Windows.Web.Http.Headers.HttpContentHeaderCollection" };
    struct headers *h = zalloc(sizeof(*h));
    if (!h) return NULL;
    void *const *vt[] = { which == H_REQUEST ? req_hdr_vtbl : which == H_RESPONSE ? resp_hdr_vtbl : cnt_hdr_vtbl, map_vtbl, iterable_vtbl, hdr_stringable_vtbl };
    const GUID *ii[] = { which == H_REQUEST ? &IID_RequestHeaders : which == H_RESPONSE ? &IID_ResponseHeaders : &IID_ContentHeaders,
                         &IID_MapSS, &IID_IterableSS, &IID_Stringable };
    obj_init(&h->m.o, K_HEADERS, names[which], headers_destroy, 4, vt, ii);
    InitializeSRWLock(&h->m.lock); h->m.l.icase = TRUE; h->which = which;
    return h;
}

/* ---- buffers --------------------------------------------------------------------------------- */
struct buffer { struct obj o; BYTE *data; UINT32 cap, len; };
static void buffer_destroy(struct obj *o) { mfree(((struct buffer *)o)->data); }
static HRESULT WINAPI buf_get_Capacity(struct itf *s, UINT32 *v) { if (!v) return E_POINTER; *v = ((struct buffer *)s->obj)->cap; return S_OK; }
static HRESULT WINAPI buf_get_Length(struct itf *s, UINT32 *v) { if (!v) return E_POINTER; *v = ((struct buffer *)s->obj)->len; return S_OK; }
static HRESULT WINAPI buf_put_Length(struct itf *s, UINT32 v)
{
    struct buffer *b = (struct buffer *)s->obj;
    if (v > b->cap) return E_INVALIDARG;
    b->len = v;
    return S_OK;
}
static HRESULT WINAPI buf_Buffer(struct itf *s, BYTE **v) { if (!v) return E_POINTER; *v = ((struct buffer *)s->obj)->data; return S_OK; }
static void *const buffer_vtbl[] = { INSPECTABLE, (void *)buf_get_Capacity, (void *)buf_get_Length, (void *)buf_put_Length };
static void *const bytes_vtbl[] = { UNKNOWN, (void *)buf_Buffer };   /* IBufferByteAccess */
static struct buffer *buffer_new(const BYTE *data, UINT32 len)
{
    struct buffer *b = zalloc(sizeof(*b));
    if (!b) return NULL;
    void *const *vt[] = { buffer_vtbl, bytes_vtbl };
    const GUID *ii[] = { &IID_Buffer, &IID_BufferByteAccess };
    obj_init(&b->o, K_BUFFER, L"Windows.Storage.Streams.Buffer", buffer_destroy, 2, vt, ii);
    if (!(b->data = zalloc(len ? len : 1))) { mfree(b); return NULL; }
    if (data && len) memcpy(b->data, data, len);
    b->cap = b->len = len;
    return b;
}
/* the bytes of any IBuffer (through IBufferByteAccess) */
static HRESULT buffer_bytes(void *ibuffer, BYTE **data, UINT32 *len, UINT32 *cap)
{
    void *access = NULL, *b = NULL;
    HRESULT hr;
    if (!ibuffer) return E_INVALIDARG;
    if (FAILED(hr = unk_qi(ibuffer, &IID_Buffer, &b))) return hr;
    if (len) VCALL(b, 7, HRESULT (WINAPI *)(void *, UINT32 *))(b, len);
    if (cap) VCALL(b, 6, HRESULT (WINAPI *)(void *, UINT32 *))(b, cap);
    if (SUCCEEDED(hr = unk_qi(ibuffer, &IID_BufferByteAccess, &access))) {
        hr = VCALL(access, 3, HRESULT (WINAPI *)(void *, BYTE **))(access, data);
        unk_release(access);
    }
    unk_release(b);
    return hr;
}

/* ---- content ----------------------------------------------------------------------------------- */
struct content { struct obj o; BYTE *data; UINT32 len; struct headers *h; };
static void content_destroy(struct obj *o)
{
    struct content *c = (struct content *)o;
    mfree(c->data); obj_release(c->h ? &c->h->m.o : NULL);
}
struct istream { struct obj o; SRWLOCK lock; struct content *c; UINT32 pos; };
static void istream_destroy(struct obj *o) { obj_release(&((struct istream *)o)->c->o); }
static HRESULT WINAPI is_ReadAsync(struct itf *s, void *buffer, UINT32 count, INT32 options, void **out)
{
    struct istream *is = (struct istream *)s->obj;
    BYTE *dst = NULL;
    UINT32 cap = 0;
    (void)options;
    HRESULT hr = buffer_bytes(buffer, &dst, NULL, &cap);
    if (FAILED(hr)) return hr;
    AcquireSRWLockExclusive(&is->lock);
    UINT32 n = is->c->len - is->pos;
    if (n > count) n = count;
    if (n > cap) n = cap;
    if (n) memcpy(dst, is->c->data + is->pos, n);
    is->pos += n;
    ReleaseSRWLockExclusive(&is->lock);
    void *b = NULL;
    if (SUCCEEDED(unk_qi(buffer, &IID_Buffer, &b))) { VCALL(b, 8, HRESULT (WINAPI *)(void *, UINT32))(b, n); unk_release(b); }   /* put_Length */
    return op_done(&IID_OpBuffer32, R_OBJ, buffer, NULL, 0, out);
}
static void *const istream_vtbl[] = { INSPECTABLE, (void *)is_ReadAsync };
static struct istream *istream_new(struct content *c)
{
    struct istream *is = zalloc(sizeof(*is));
    if (!is) return NULL;
    void *const *vt[] = { istream_vtbl, closable_vtbl };
    const GUID *ii[] = { &IID_InputStream, &IID_Closable };
    obj_init(&is->o, K_ISTREAM, L"Windows.Storage.Streams.IInputStream", istream_destroy, 2, vt, ii);
    InitializeSRWLock(&is->lock);
    is->c = c; obj_addref(&c->o);
    return is;
}

static HRESULT WINAPI cnt_get_Headers(struct itf *s, void **out)
{
    struct content *c = (struct content *)s->obj;
    if (!out) return E_POINTER;
    obj_addref(&c->h->m.o);
    ret_itf(&c->h->m.o, out);
    return S_OK;
}
static HRESULT WINAPI cnt_BufferAllAsync(struct itf *s, void **out)
{
    return op_done(&IID_OpU64U64, R_U64, NULL, NULL, ((struct content *)s->obj)->len, out);
}
static HRESULT WINAPI cnt_ReadAsBufferAsync(struct itf *s, void **out)
{
    struct content *c = (struct content *)s->obj;
    struct buffer *b = buffer_new(c->data, c->len);
    if (!b) return E_OUTOFMEMORY;
    HRESULT hr = op_done(&IID_OpBuffer64, R_OBJ, itf0(&b->o), NULL, 0, out);
    obj_release(&b->o);
    return hr;
}
static HRESULT WINAPI cnt_ReadAsInputStreamAsync(struct itf *s, void **out)
{
    struct istream *is = istream_new((struct content *)s->obj);
    if (!is) return E_OUTOFMEMORY;
    HRESULT hr = op_done(&IID_OpInputStream64, R_OBJ, itf0(&is->o), NULL, 0, out);
    obj_release(&is->o);
    return hr;
}
static HRESULT WINAPI cnt_ReadAsStringAsync(struct itf *s, void **out)
{
    struct content *c = (struct content *)s->obj;
    WCHAR *ct = hdr_get(c->h, L"Content-Type"), *cs = ct ? param_value(ct, L"charset") : NULL, *text;
    if (cs && (!_wcsicmp(cs, L"utf-16") || !_wcsicmp(cs, L"utf-16le") || !_wcsicmp(cs, L"unicode")))
        text = wdup_n((const WCHAR *)c->data, c->len / sizeof(WCHAR));
    else
        text = wide_of_utf8(c->data, c->len);
    mfree(ct); mfree(cs);
    if (!text) return E_OUTOFMEMORY;
    HRESULT hr = op_done(&IID_OpString64, R_STR, NULL, text, 0, out);
    mfree(text);
    return hr;
}
static HRESULT WINAPI cnt_TryComputeLength(struct itf *s, UINT64 *len, boolean8 *ok)
{
    if (!len || !ok) return E_POINTER;
    *len = ((struct content *)s->obj)->len; *ok = 1;
    return S_OK;
}
NOTIMPL(cnt_WriteToStreamAsync, "IHttpContent.WriteToStreamAsync")
static void *const content_vtbl[] = { INSPECTABLE, (void *)cnt_get_Headers, (void *)cnt_BufferAllAsync, (void *)cnt_ReadAsBufferAsync,
    (void *)cnt_ReadAsInputStreamAsync, (void *)cnt_ReadAsStringAsync, (void *)cnt_TryComputeLength, (void *)cnt_WriteToStreamAsync };
/* takes ownership of data */
static struct content *content_new(const WCHAR *cls, BYTE *data, UINT32 len)
{
    struct content *c = zalloc(sizeof(*c));
    if (!c || !(c->h = headers_new(H_CONTENT))) { mfree(c); mfree(data); return NULL; }
    void *const *vt[] = { content_vtbl, closable_vtbl, stringable_vtbl };
    const GUID *ii[] = { &IID_Content, &IID_Closable, &IID_Stringable };
    obj_init(&c->o, K_CONTENT, cls, content_destroy, 3, vt, ii);
    c->data = data; c->len = data ? len : 0;
    return c;
}

/* ---- DataReader -------------------------------------------------------------------------------- */
#define MAX_QUEUE (1u << 30)
struct reader { struct obj o; SRWLOCK lock; void *stream; BYTE *q; UINT32 head, tail, cap; INT32 encoding, order, options; };
static void reader_destroy(struct obj *o) { struct reader *r = (struct reader *)o; unk_release(r->stream); mfree(r->q); }
static BOOL reader_push(struct reader *r, const BYTE *p, UINT32 n)
{
    if ((UINT64)r->tail - r->head + n > MAX_QUEUE) return FALSE;
    if (r->head) { memmove(r->q, r->q + r->head, r->tail - r->head); r->tail -= r->head; r->head = 0; }
    if (r->tail + n > r->cap) {
        UINT32 cap = r->cap ? r->cap : 4096;
        while (cap < r->tail + n) cap *= 2;
        BYTE *q = grow(r->q, cap);
        if (!q) return FALSE;
        r->q = q; r->cap = cap;
    }
    memcpy(r->q + r->tail, p, n); r->tail += n;
    return TRUE;
}
static HRESULT reader_take(struct reader *r, void *dst, UINT32 n, BOOL ordered)
{
    AcquireSRWLockExclusive(&r->lock);
    if (r->tail - r->head < n) { ReleaseSRWLockExclusive(&r->lock); return E_BOUNDS_; }
    BYTE *src = r->q + r->head;
    if (ordered && r->order == 1 /* BigEndian */) for (UINT32 i = 0; i < n; i++) ((BYTE *)dst)[i] = src[n - 1 - i];
    else if (n) memcpy(dst, src, n);
    r->head += n;
    ReleaseSRWLockExclusive(&r->lock);
    return S_OK;
}
static HRESULT WINAPI rd_get_Unconsumed(struct itf *s, UINT32 *v)
{
    struct reader *r = (struct reader *)s->obj;
    if (!v) return E_POINTER;
    AcquireSRWLockShared(&r->lock); *v = r->tail - r->head; ReleaseSRWLockShared(&r->lock);
    return S_OK;
}
static HRESULT WINAPI rd_get_Encoding(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct reader *)s->obj)->encoding; return S_OK; }
static HRESULT WINAPI rd_put_Encoding(struct itf *s, INT32 v) { ((struct reader *)s->obj)->encoding = v; return S_OK; }
static HRESULT WINAPI rd_get_Order(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct reader *)s->obj)->order; return S_OK; }
static HRESULT WINAPI rd_put_Order(struct itf *s, INT32 v) { ((struct reader *)s->obj)->order = v; return S_OK; }
static HRESULT WINAPI rd_get_Options(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct reader *)s->obj)->options; return S_OK; }
static HRESULT WINAPI rd_put_Options(struct itf *s, INT32 v) { ((struct reader *)s->obj)->options = v; return S_OK; }
static HRESULT WINAPI rd_ReadByte(struct itf *s, BYTE *v) { return v ? reader_take((struct reader *)s->obj, v, 1, FALSE) : E_POINTER; }
static HRESULT WINAPI rd_ReadBytes(struct itf *s, UINT32 n, BYTE *v) { return (v || !n) ? reader_take((struct reader *)s->obj, v, n, FALSE) : E_POINTER; }
static HRESULT WINAPI rd_ReadBuffer(struct itf *s, UINT32 n, void **out)
{
    struct buffer *b = buffer_new(NULL, n);
    if (!out) { obj_release(b ? &b->o : NULL); return E_POINTER; }
    *out = NULL;
    if (!b) return E_OUTOFMEMORY;
    HRESULT hr = reader_take((struct reader *)s->obj, b->data, n, FALSE);
    if (FAILED(hr)) { obj_release(&b->o); return hr; }
    ret_itf(&b->o, out);
    return S_OK;
}
static HRESULT WINAPI rd_ReadBoolean(struct itf *s, boolean8 *v)
{
    BYTE b = 0;
    HRESULT hr = v ? reader_take((struct reader *)s->obj, &b, 1, FALSE) : E_POINTER;
    if (SUCCEEDED(hr)) *v = b != 0;
    return hr;
}
static HRESULT WINAPI rd_ReadGuid(struct itf *s, GUID *v)
{
    struct reader *r = (struct reader *)s->obj;
    if (!v) return E_POINTER;
    HRESULT hr = reader_take(r, &v->Data1, 4, TRUE);
    if (SUCCEEDED(hr)) hr = reader_take(r, &v->Data2, 2, TRUE);
    if (SUCCEEDED(hr)) hr = reader_take(r, &v->Data3, 2, TRUE);
    if (SUCCEEDED(hr)) hr = reader_take(r, v->Data4, 8, FALSE);
    return hr;
}
static HRESULT WINAPI rd_Read2(struct itf *s, void *v) { return v ? reader_take((struct reader *)s->obj, v, 2, TRUE) : E_POINTER; }
static HRESULT WINAPI rd_Read4(struct itf *s, void *v) { return v ? reader_take((struct reader *)s->obj, v, 4, TRUE) : E_POINTER; }
static HRESULT WINAPI rd_Read8(struct itf *s, void *v) { return v ? reader_take((struct reader *)s->obj, v, 8, TRUE) : E_POINTER; }
static HRESULT WINAPI rd_ReadString(struct itf *s, UINT32 count, HSTRING *out)
{
    struct reader *r = (struct reader *)s->obj;
    UINT32 unit = r->encoding == 0 /* Utf8 */ ? 1 : 2;
    BYTE *raw = zalloc((SIZE_T)count * unit + 2);
    WCHAR *text = NULL;
    if (!out) { mfree(raw); return E_POINTER; }
    *out = NULL;
    if (!raw) return E_OUTOFMEMORY;
    HRESULT hr = reader_take(r, raw, count * unit, FALSE);
    if (SUCCEEDED(hr)) {
        if (unit == 1) text = wide_of_utf8(raw, count);
        else {
            if (r->encoding == 2 /* Utf16BE */) for (UINT32 i = 0; i < count; i++) { BYTE t = raw[2 * i]; raw[2 * i] = raw[2 * i + 1]; raw[2 * i + 1] = t; }
            text = wdup_n((const WCHAR *)raw, count);
        }
        hr = text ? to_hstring(text, out) : E_OUTOFMEMORY;
    }
    mfree(raw); mfree(text);
    return hr;
}
static HRESULT WINAPI rd_LoadAsync(struct itf *s, UINT32 count, void **out)
{
    struct reader *r = (struct reader *)s->obj;
    struct istream *is = (struct istream *)ours(r->stream, K_ISTREAM);
    UINT32 n = 0;
    if (r->stream && !is) return notimpl("DataReader.LoadAsync on a stream from elsewhere");
    if (is) {
        AcquireSRWLockExclusive(&is->lock);
        n = is->c->len - is->pos;
        if (n > count) n = count;
        AcquireSRWLockExclusive(&r->lock);
        if (n && !reader_push(r, is->c->data + is->pos, n)) n = 0;
        ReleaseSRWLockExclusive(&r->lock);
        is->pos += n;
        ReleaseSRWLockExclusive(&is->lock);
    }
    return op_done(&IID_OpU32, R_U32, NULL, NULL, n, out);
}
NOTIMPL(rd_DetachBuffer, "DataReader.DetachBuffer")
NOTIMPL(rd_DetachStream, "DataReader.DetachStream")
static HRESULT WINAPI rd_Close(struct itf *s)
{
    struct reader *r = (struct reader *)s->obj;
    void *stream = InterlockedExchangePointer(&r->stream, NULL);
    unk_release(stream);
    return S_OK;
}
static void *const reader_vtbl[] = { INSPECTABLE,
    (void *)rd_get_Unconsumed, (void *)rd_get_Encoding, (void *)rd_put_Encoding, (void *)rd_get_Order, (void *)rd_put_Order,
    (void *)rd_get_Options, (void *)rd_put_Options, (void *)rd_ReadByte, (void *)rd_ReadBytes, (void *)rd_ReadBuffer,
    (void *)rd_ReadBoolean, (void *)rd_ReadGuid, (void *)rd_Read2 /* Int16 */, (void *)rd_Read4 /* Int32 */, (void *)rd_Read8 /* Int64 */,
    (void *)rd_Read2 /* UInt16 */, (void *)rd_Read4 /* UInt32 */, (void *)rd_Read8 /* UInt64 */, (void *)rd_Read4 /* Single */,
    (void *)rd_Read8 /* Double */, (void *)rd_ReadString, (void *)rd_Read8 /* DateTime */, (void *)rd_Read8 /* TimeSpan */,
    (void *)rd_LoadAsync, (void *)rd_DetachBuffer, (void *)rd_DetachStream };
static void *const reader_closable_vtbl[] = { INSPECTABLE, (void *)rd_Close };
static struct reader *reader_new(void *stream)
{
    struct reader *r = zalloc(sizeof(*r));
    if (!r) return NULL;
    void *const *vt[] = { reader_vtbl, reader_closable_vtbl };
    const GUID *ii[] = { &IID_DataReader, &IID_Closable };
    obj_init(&r->o, K_READER, L"Windows.Storage.Streams.DataReader", reader_destroy, 2, vt, ii);
    InitializeSRWLock(&r->lock);
    r->stream = stream; unk_addref(stream);
    return r;
}

/* ---- HttpMethod -------------------------------------------------------------------------------- */
struct method { struct obj o; WCHAR name[24]; };
static HRESULT WINAPI m_get_Method(struct itf *s, HSTRING *v) { return to_hstring(((struct method *)s->obj)->name, v); }
static void *const method_vtbl[] = { INSPECTABLE, (void *)m_get_Method };
static void *const method_stringable_vtbl[] = { INSPECTABLE, (void *)m_get_Method };
static struct obj *method_new(const WCHAR *name)
{
    struct method *m = zalloc(sizeof(*m));
    if (!m) return NULL;
    void *const *vt[] = { method_vtbl, method_stringable_vtbl };
    const GUID *ii[] = { &IID_Method, &IID_Stringable };
    obj_init(&m->o, K_METHOD, L"Windows.Web.Http.HttpMethod", NULL, 2, vt, ii);
    lstrcpynW(m->name, name, sizeof(m->name) / sizeof(WCHAR));
    return &m->o;
}
static WCHAR *method_string(void *method)   /* IHttpMethod::get_Method */
{
    HSTRING h = NULL;
    void *m = NULL;
    WCHAR *s = NULL;
    if (method && SUCCEEDED(unk_qi(method, &IID_Method, &m))) {
        if (SUCCEEDED(VCALL(m, 6, HRESULT (WINAPI *)(void *, HSTRING *))(m, &h))) s = from_hstring(h);
        if (h) p_WindowsDeleteString(h);
        unk_release(m);
    }
    return s;
}

/* ---- HttpRequestMessage ------------------------------------------------------------------------ */
struct request { struct obj o; SRWLOCK lock; IUnknown *method, *uri, *content; struct headers *h; struct map *props; };
static void request_destroy(struct obj *o)
{
    struct request *r = (struct request *)o;
    unk_release(r->method); unk_release(r->uri); unk_release(r->content);
    obj_release(r->h ? &r->h->m.o : NULL); obj_release(r->props ? &r->props->o : NULL);
}
static HRESULT get_member(SRWLOCK *lock, IUnknown **slot, void **out)
{
    if (!out) return E_POINTER;
    AcquireSRWLockShared(lock); *out = *slot; unk_addref(*out); ReleaseSRWLockShared(lock);
    return S_OK;
}
static HRESULT put_member(SRWLOCK *lock, IUnknown **slot, void *v)
{
    AcquireSRWLockExclusive(lock); set_unk(slot, v); ReleaseSRWLockExclusive(lock);
    return S_OK;
}
static HRESULT WINAPI rq_get_Content(struct itf *s, void **v) { struct request *r = (struct request *)s->obj; return get_member(&r->lock, &r->content, v); }
static HRESULT WINAPI rq_put_Content(struct itf *s, void *v) { struct request *r = (struct request *)s->obj; return put_member(&r->lock, &r->content, v); }
static HRESULT WINAPI rq_get_Headers(struct itf *s, void **v)
{
    struct request *r = (struct request *)s->obj;
    if (!v) return E_POINTER;
    obj_addref(&r->h->m.o); ret_itf(&r->h->m.o, v);
    return S_OK;
}
static HRESULT WINAPI rq_get_Method(struct itf *s, void **v) { struct request *r = (struct request *)s->obj; return get_member(&r->lock, &r->method, v); }
static HRESULT WINAPI rq_put_Method(struct itf *s, void *v) { struct request *r = (struct request *)s->obj; return v ? put_member(&r->lock, &r->method, v) : E_INVALIDARG; }
static HRESULT WINAPI rq_get_Properties(struct itf *s, void **v)
{
    struct request *r = (struct request *)s->obj;
    if (!v) return E_POINTER;
    obj_addref(&r->props->o); ret_itf(&r->props->o, v);
    return S_OK;
}
static HRESULT WINAPI rq_get_RequestUri(struct itf *s, void **v) { struct request *r = (struct request *)s->obj; return get_member(&r->lock, &r->uri, v); }
static HRESULT WINAPI rq_put_RequestUri(struct itf *s, void *v) { struct request *r = (struct request *)s->obj; return put_member(&r->lock, &r->uri, v); }
NOTIMPL(rq_get_TransportInformation, "HttpRequestMessage.TransportInformation")
static HRESULT WINAPI rq_get_PrivacyAnnotation(struct itf *s, HSTRING *v) { (void)s; if (!v) return E_POINTER; *v = NULL; return S_OK; }
static HRESULT WINAPI rq_put_PrivacyAnnotation(struct itf *s, HSTRING v) { (void)s; (void)v; return S_OK; }
static HRESULT WINAPI rq_ToString(struct itf *s, HSTRING *out)
{
    struct request *r = (struct request *)s->obj;
    AcquireSRWLockShared(&r->lock);
    WCHAR *m = method_string(r->method), *u = uri_string(r->uri);
    ReleaseSRWLockShared(&r->lock);
    WCHAR *t = wconcat(m ? m : L"", L" ", u ? u : L"");
    HRESULT hr = t ? to_hstring(t, out) : E_OUTOFMEMORY;
    mfree(m); mfree(u); mfree(t);
    return hr;
}
static void *const request_vtbl[] = { INSPECTABLE, (void *)rq_get_Content, (void *)rq_put_Content, (void *)rq_get_Headers, (void *)rq_get_Method,
    (void *)rq_put_Method, (void *)rq_get_Properties, (void *)rq_get_RequestUri, (void *)rq_put_RequestUri, (void *)rq_get_TransportInformation };
static void *const request2_vtbl[] = { INSPECTABLE, (void *)rq_get_PrivacyAnnotation, (void *)rq_put_PrivacyAnnotation };
static void *const request_stringable_vtbl[] = { INSPECTABLE, (void *)rq_ToString };
static struct request *request_new(void *method, void *uri)
{
    struct request *r = zalloc(sizeof(*r));
    if (!r) return NULL;
    void *const *vt[] = { request_vtbl, request2_vtbl, closable_vtbl, request_stringable_vtbl };
    const GUID *ii[] = { &IID_Request, &IID_Request2, &IID_Closable, &IID_Stringable };
    obj_init(&r->o, K_REQUEST, L"Windows.Web.Http.HttpRequestMessage", request_destroy, 4, vt, ii);
    InitializeSRWLock(&r->lock);
    if (!(r->h = headers_new(H_REQUEST)) || !(r->props = propmap_new())) { obj_release(&r->o); return NULL; }
    if (method) { r->method = method; unk_addref(method); }
    else { struct obj *m = method_new(L"GET"); r->method = m ? itf0(m) : NULL; }
    r->uri = uri; unk_addref(uri);
    return r;
}

/* ---- HttpResponseMessage ----------------------------------------------------------------------- */
struct response { struct obj o; SRWLOCK lock; IUnknown *content, *request; struct headers *h; WCHAR *reason; INT32 status, version, source; };
static void response_destroy(struct obj *o)
{
    struct response *r = (struct response *)o;
    unk_release(r->content); unk_release(r->request); obj_release(r->h ? &r->h->m.o : NULL); mfree(r->reason);
}
static HRESULT WINAPI rs_get_Content(struct itf *s, void **v) { struct response *r = (struct response *)s->obj; return get_member(&r->lock, &r->content, v); }
static HRESULT WINAPI rs_put_Content(struct itf *s, void *v) { struct response *r = (struct response *)s->obj; return put_member(&r->lock, &r->content, v); }
static HRESULT WINAPI rs_get_Headers(struct itf *s, void **v)
{
    struct response *r = (struct response *)s->obj;
    if (!v) return E_POINTER;
    obj_addref(&r->h->m.o); ret_itf(&r->h->m.o, v);
    return S_OK;
}
static HRESULT WINAPI rs_get_IsSuccess(struct itf *s, boolean8 *v)
{
    struct response *r = (struct response *)s->obj;
    if (!v) return E_POINTER;
    *v = r->status >= 200 && r->status <= 299;
    return S_OK;
}
static HRESULT WINAPI rs_get_ReasonPhrase(struct itf *s, HSTRING *v)
{
    struct response *r = (struct response *)s->obj;
    AcquireSRWLockShared(&r->lock); HRESULT hr = to_hstring(r->reason, v); ReleaseSRWLockShared(&r->lock);
    return hr;
}
static HRESULT WINAPI rs_put_ReasonPhrase(struct itf *s, HSTRING v)
{
    struct response *r = (struct response *)s->obj;
    WCHAR *t = from_hstring(v);
    AcquireSRWLockExclusive(&r->lock); mfree(r->reason); r->reason = t; ReleaseSRWLockExclusive(&r->lock);
    return S_OK;
}
static HRESULT WINAPI rs_get_RequestMessage(struct itf *s, void **v) { struct response *r = (struct response *)s->obj; return get_member(&r->lock, &r->request, v); }
static HRESULT WINAPI rs_put_RequestMessage(struct itf *s, void *v) { struct response *r = (struct response *)s->obj; return put_member(&r->lock, &r->request, v); }
static HRESULT WINAPI rs_get_Source(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct response *)s->obj)->source; return S_OK; }
static HRESULT WINAPI rs_put_Source(struct itf *s, INT32 v) { ((struct response *)s->obj)->source = v; return S_OK; }
static HRESULT WINAPI rs_get_StatusCode(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct response *)s->obj)->status; return S_OK; }
static HRESULT WINAPI rs_put_StatusCode(struct itf *s, INT32 v) { ((struct response *)s->obj)->status = v; return S_OK; }
static HRESULT WINAPI rs_get_Version(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct response *)s->obj)->version; return S_OK; }
static HRESULT WINAPI rs_put_Version(struct itf *s, INT32 v) { ((struct response *)s->obj)->version = v; return S_OK; }
static HRESULT WINAPI rs_EnsureSuccessStatusCode(struct itf *s, void **out)
{
    struct response *r = (struct response *)s->obj;
    if (!out) return E_POINTER;
    *out = NULL;
    if (r->status < 200 || r->status > 299) return MAKE_HRESULT(SEVERITY_ERROR, 0x19, r->status);   /* HTTP_E_STATUS_* */
    obj_addref(&r->o); ret_itf(&r->o, out);
    return S_OK;
}
static HRESULT WINAPI rs_ToString(struct itf *s, HSTRING *out)
{
    WCHAR buf[64];
    wsprintfW(buf, L"StatusCode: %d", ((struct response *)s->obj)->status);
    return to_hstring(buf, out);
}
static void *const response_vtbl[] = { INSPECTABLE, (void *)rs_get_Content, (void *)rs_put_Content, (void *)rs_get_Headers, (void *)rs_get_IsSuccess,
    (void *)rs_get_ReasonPhrase, (void *)rs_put_ReasonPhrase, (void *)rs_get_RequestMessage, (void *)rs_put_RequestMessage,
    (void *)rs_get_Source, (void *)rs_put_Source, (void *)rs_get_StatusCode, (void *)rs_put_StatusCode,
    (void *)rs_get_Version, (void *)rs_put_Version, (void *)rs_EnsureSuccessStatusCode };
static void *const response_stringable_vtbl[] = { INSPECTABLE, (void *)rs_ToString };
static struct response *response_new(INT32 status)
{
    struct response *r = zalloc(sizeof(*r));
    struct content *c = NULL;
    if (!r) return NULL;
    void *const *vt[] = { response_vtbl, closable_vtbl, response_stringable_vtbl };
    const GUID *ii[] = { &IID_Response, &IID_Closable, &IID_Stringable };
    obj_init(&r->o, K_RESPONSE, L"Windows.Web.Http.HttpResponseMessage", response_destroy, 3, vt, ii);
    InitializeSRWLock(&r->lock);
    r->status = status; r->version = 2 /* Http11 */;
    if (!(r->h = headers_new(H_RESPONSE)) || !(c = content_new(L"Windows.Web.Http.HttpStreamContent", NULL, 0))) { obj_release(&r->o); return NULL; }
    r->content = itf0(&c->o);   /* an empty body, as on Windows */
    return r;
}

/* ---- sending a request on WinHTTP ---------------------------------------------------------------- */
#define MAX_BODY (512u << 20)   /* response bodies are held in memory whole */
struct send {
    SRWLOCK lock; BOOL canceled; HINTERNET connect, req;
    struct op *op; struct request *request;
    WCHAR *method, *url; struct kvlist headers; BYTE *body; UINT32 body_len; BOOL redirects, cookies;
};
static void send_free(struct send *x)
{
    obj_release(&x->op->o); obj_release(&x->request->o);
    mfree(x->method); mfree(x->url); kv_clear(&x->headers); mfree(x->body); mfree(x);
}
static void send_close(struct send *x)   /* WinHttpCloseHandle makes a blocked call on the request return */
{
    AcquireSRWLockExclusive(&x->lock);
    if (x->req) p_WinHttpCloseHandle(x->req);
    if (x->connect) p_WinHttpCloseHandle(x->connect);
    x->req = x->connect = NULL;
    ReleaseSRWLockExclusive(&x->lock);
}
static void send_cancel(void *ctx) { struct send *x = ctx; x->canceled = TRUE; send_close(x); }   /* under the op's lock */
static HRESULT last_error(void) { DWORD e = GetLastError(); return e ? HRESULT_FROM_WIN32(e) : E_FAIL; }
static WCHAR *query_string(HINTERNET req, DWORD info)
{
    DWORD len = 0;
    p_WinHttpQueryHeaders(req, info, WINHTTP_HEADER_NAME_BY_INDEX, NULL, &len, WINHTTP_NO_HEADER_INDEX);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || !len) return NULL;
    WCHAR *buf = zalloc(len + sizeof(WCHAR));
    if (buf && !p_WinHttpQueryHeaders(req, info, WINHTTP_HEADER_NAME_BY_INDEX, buf, &len, WINHTTP_NO_HEADER_INDEX)) { mfree(buf); return NULL; }
    return buf;
}
static BOOL is_content_header(const WCHAR *name)
{
    static const WCHAR *const names[] = { L"Content-Type", L"Content-Length", L"Content-Encoding", L"Content-Language", L"Content-Location",
                                          L"Content-MD5", L"Content-Range", L"Content-Disposition", L"Expires", L"Last-Modified" };
    for (size_t i = 0; i < sizeof(names) / sizeof(names[0]); i++) if (!_wcsicmp(name, names[i])) return TRUE;
    return FALSE;
}
/* the response headers after the status line, split between the response and its content */
static void parse_headers(const WCHAR *raw, struct headers *rh, struct headers *ch)
{
    const WCHAR *p = raw ? wcsstr(raw, L"\r\n") : NULL;
    while (p && p[0]) {
        p += 2;
        const WCHAR *e = wcsstr(p, L"\r\n");
        if (!e) e = p + wcslen(p);
        if (e == p) break;
        const WCHAR *colon = wmemchr(p, ':', e - p);
        if (colon) {
            WCHAR *name = wtrim_dup(p, colon - p), *value = wtrim_dup(colon + 1, e - colon - 1);
            if (name && value && name[0]) hdr_set(is_content_header(name) ? ch : rh, name, value, TRUE);
            mfree(name); mfree(value);
        }
        p = e;
    }
}
static void log_request(struct send *x, const char *result)
{
    WCHAR url[200];
    lstrcpynW(url, x->url, sizeof(url) / sizeof(WCHAR));
    WCHAR *q = wcspbrk(url, L"?#");   /* no query strings in the log */
    if (q) *q = 0;
    wh_log("ms365 webhttp: %ls %ls -> %s", x->method, url, result);
}
/* one request to x->url, up to the response headers; the handles go to x so a cancel can close them */
static HRESULT send_open(struct send *x, HINTERNET *out)
{
    URL_COMPONENTS uc;
    WCHAR *host = NULL, *path = NULL;
    HRESULT hr = S_OK;
    *out = NULL;
    memset(&uc, 0, sizeof(uc)); uc.dwStructSize = sizeof(uc);
    uc.dwSchemeLength = uc.dwHostNameLength = uc.dwUrlPathLength = uc.dwExtraInfoLength = (DWORD)-1;
    if (!p_WinHttpCrackUrl(x->url, 0, 0, &uc)) return last_error();
    host = wdup_n(uc.lpszHostName, uc.dwHostNameLength);
    path = wdup_n(uc.lpszUrlPath, uc.dwUrlPathLength + uc.dwExtraInfoLength);   /* path and query are adjacent */
    if (!host || !path) { hr = E_OUTOFMEMORY; goto done; }
    HINTERNET connect = p_WinHttpConnect(g_session, host, uc.nPort, 0);
    HINTERNET req = connect ? p_WinHttpOpenRequest(connect, x->method, path, NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                                                   uc.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0) : NULL;
    if (!req) hr = last_error();
    AcquireSRWLockExclusive(&x->lock);
    x->connect = connect; x->req = req;
    if (x->canceled) hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    ReleaseSRWLockExclusive(&x->lock);
    if (FAILED(hr)) goto done;
    /* WinHTTP's own redirects would send the caller's headers, Authorization included, to whatever
     * host the server names; redirects are followed in send_run instead */
    DWORD v = WINHTTP_DISABLE_REDIRECTS | (x->cookies ? 0 : WINHTTP_DISABLE_COOKIES);
    if (!p_WinHttpSetOption(req, WINHTTP_OPTION_DISABLE_FEATURE, &v, sizeof(v))) { hr = last_error(); goto done; }
    for (UINT32 i = 0; i < x->headers.n; i++) {
        const struct kv *e = &x->headers.v[i];
        if (!_wcsicmp(e->key, L"Host") || !_wcsicmp(e->key, L"Content-Length")) continue;   /* WinHTTP's own */
        if (!is_token(e->key) || !is_field_value(e->str)) continue;
        WCHAR *line = wconcat(e->key, L": ", e->str ? e->str : L"");
        if (line && !p_WinHttpAddRequestHeaders(req, line, (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE))
            wh_log("ms365 webhttp: header %.100ls refused (%lu)", e->key, GetLastError());
        mfree(line);
    }
    if (!p_WinHttpSendRequest(req, WINHTTP_NO_ADDITIONAL_HEADERS, 0, x->body_len ? x->body : WINHTTP_NO_REQUEST_DATA,
                              x->body_len, x->body_len, 0) ||
        !p_WinHttpReceiveResponse(req, NULL)) { hr = last_error(); goto done; }
    *out = req;
done:
    mfree(host); mfree(path);
    return hr;
}
static BOOL same_origin(const WCHAR *a, const WCHAR *b)
{
    URL_COMPONENTS ua, ub;
    memset(&ua, 0, sizeof(ua)); ua.dwStructSize = sizeof(ua); ua.dwHostNameLength = (DWORD)-1;
    ub = ua;
    if (!p_WinHttpCrackUrl(a, 0, 0, &ua) || !p_WinHttpCrackUrl(b, 0, 0, &ub)) return FALSE;
    return ua.nScheme == ub.nScheme && ua.nPort == ub.nPort && ua.dwHostNameLength == ub.dwHostNameLength &&
           !_wcsnicmp(ua.lpszHostName, ub.lpszHostName, ua.dwHostNameLength);
}
/* AllowAutoRedirect: point x at a redirect's target, as browsers do. Credentials are dropped once
 * the origin changes, 301 and 302 turn a POST into a GET and 303 turns anything but HEAD into one,
 * and HTTPS never redirects to HTTP (WinHTTP's default policy). FALSE: return the response as it is. */
static BOOL follow_redirect(struct send *x, HINTERNET req, DWORD status)
{
    static const WCHAR *const credentials[] = { L"Authorization", L"Proxy-Authorization", L"Cookie" };
    if (status != 301 && status != 302 && status != 303 && status != 307 && status != 308) return FALSE;
    WCHAR *loc = query_string(req, WINHTTP_QUERY_LOCATION);
    WCHAR *next = loc ? resolve_location(x->url, loc) : NULL;
    mfree(loc);
    if (!next) return FALSE;
    if (_wcsnicmp(next, L"https://", 8) && (!_wcsnicmp(x->url, L"https:", 6) || _wcsnicmp(next, L"http://", 7))) { mfree(next); return FALSE; }
    if (status == 303 ? _wcsicmp(x->method, L"HEAD") != 0 : status <= 302 && !_wcsicmp(x->method, L"POST")) {
        WCHAR *get = wdup(L"GET");
        if (!get) { mfree(next); return FALSE; }
        mfree(x->method); x->method = get;
        mfree(x->body); x->body = NULL; x->body_len = 0;
        for (UINT32 i = x->headers.n; i-- > 0;) if (is_content_header(x->headers.v[i].key)) kv_remove(&x->headers, x->headers.v[i].key);
    }
    if (!same_origin(x->url, next))
        for (size_t i = 0; i < sizeof(credentials) / sizeof(credentials[0]); i++) kv_remove(&x->headers, credentials[i]);
    mfree(x->url); x->url = next;
    return TRUE;
}
#define MAX_REDIRECTS 10   /* WinHTTP's default */
static HRESULT send_run(struct send *x, struct response **out)
{
    WCHAR *raw = NULL, *reason = NULL, *version = NULL;
    BYTE *body = NULL;
    UINT32 body_len = 0, body_cap = 0;
    HINTERNET req;
    DWORD status;
    HRESULT hr = S_OK;
    *out = NULL;
    if (!g_session) return HRESULT_FROM_WIN32(ERROR_WINHTTP_INTERNAL_ERROR);
    for (UINT32 hops = 0;; hops++) {
        if (FAILED(hr = send_open(x, &req))) goto done;
        DWORD sz = sizeof(status);
        status = 0;
        p_WinHttpQueryHeaders(req, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &sz, WINHTTP_NO_HEADER_INDEX);
        /* react-native-windows turns AllowAutoRedirect off and follows redirects itself (RedirectHttpFilter) */
        if (!x->redirects || !follow_redirect(x, req, status)) break;
        if (hops == MAX_REDIRECTS) { hr = HRESULT_FROM_WIN32(ERROR_WINHTTP_REDIRECT_FAILED); goto done; }
        send_close(x);
    }
    reason = query_string(req, WINHTTP_QUERY_STATUS_TEXT);
    version = query_string(req, WINHTTP_QUERY_VERSION);
    raw = query_string(req, WINHTTP_QUERY_RAW_HEADERS_CRLF);
    for (;;) {
        DWORD avail = 0, got = 0;
        if (!p_WinHttpQueryDataAvailable(req, &avail)) { hr = last_error(); goto done; }
        if (!avail) break;
        if ((UINT64)body_len + avail > MAX_BODY) { hr = E_OUTOFMEMORY; goto done; }
        if (body_len + avail > body_cap) {
            UINT32 cap = body_cap ? body_cap : 16384;
            while (cap < body_len + avail) cap *= 2;   /* at most 2 * MAX_BODY */
            BYTE *b = grow(body, cap);
            if (!b) { hr = E_OUTOFMEMORY; goto done; }
            body = b; body_cap = cap;
        }
        if (!p_WinHttpReadData(req, body + body_len, avail, &got)) { hr = last_error(); goto done; }
        if (!got) break;
        body_len += got;
    }
    struct response *r = response_new((INT32)status);
    struct content *c = NULL;
    if (r) { c = content_new(L"Windows.Web.Http.HttpStreamContent", body, body_len); body = NULL; }   /* takes body */
    if (!c) { obj_release(r ? &r->o : NULL); hr = E_OUTOFMEMORY; goto done; }
    set_unk(&r->content, itf0(&c->o)); obj_release(&c->o);
    r->reason = reason; reason = NULL;
    r->version = !version ? 2 : !wcscmp(version, L"HTTP/1.0") ? 1 : !wcsncmp(version, L"HTTP/2", 6) ? 3 : 2;
    r->source = 2 /* Network */;
    r->h->base = wdup(x->url);
    parse_headers(raw, r->h, c->h);
    r->request = itf0(&x->request->o); obj_addref(&x->request->o);
    *out = r;
    { char res[48]; wsprintfA(res, "%lu, %u bytes", status, c->len); log_request(x, res); }
done:
    mfree(raw); mfree(reason); mfree(version); mfree(body);
    return hr;
}
static DWORD WINAPI send_thread(void *ctx)
{
    struct send *x = ctx;
    struct response *r = NULL;
    HRESULT hr = send_run(x, &r);
    AcquireSRWLockExclusive(&x->op->lock);   /* no more cancellation through x from here on */
    x->op->cancel = NULL; x->op->cancel_ctx = NULL;
    ReleaseSRWLockExclusive(&x->op->lock);
    send_close(x);
    if (x->canceled) { op_finish(x->op, ST_CANCELED, S_OK, NULL, NULL, 0); log_request(x, "canceled"); }
    else if (FAILED(hr)) { char res[32]; wsprintfA(res, "error 0x%08lx", hr); log_request(x, res); op_finish(x->op, ST_ERROR, hr, NULL, NULL, 0); }
    else op_finish(x->op, ST_COMPLETED, S_OK, itf0(&r->o), NULL, 0);
    obj_release(r ? &r->o : NULL);
    send_free(x);
    return 0;
}
static HRESULT start_send(void *request, BOOL redirects, BOOL cookies, void **out)
{
    struct request *rq = (struct request *)ours(request, K_REQUEST);
    struct content *c = NULL;
    if (!out) return E_POINTER;
    *out = NULL;
    if (!rq) return notimpl("SendRequestAsync with a request from elsewhere");
    struct send *x = zalloc(sizeof(*x));
    if (!x) return E_OUTOFMEMORY;
    InitializeSRWLock(&x->lock);
    x->redirects = redirects; x->cookies = cookies; x->headers.icase = TRUE;
    AcquireSRWLockShared(&rq->lock);
    x->method = method_string(rq->method);
    x->url = uri_string(rq->uri);
    if (rq->content && !(c = (struct content *)ours(rq->content, K_CONTENT))) {
        ReleaseSRWLockShared(&rq->lock);
        mfree(x->method); mfree(x->url); mfree(x);
        return notimpl("request content other than HttpStringContent/HttpBufferContent");
    }
    if (c) obj_addref(&c->o);
    ReleaseSRWLockShared(&rq->lock);
    AcquireSRWLockShared(&rq->h->m.lock); kv_copy(&x->headers, &rq->h->m.l); ReleaseSRWLockShared(&rq->h->m.lock);
    if (c) {
        AcquireSRWLockShared(&c->h->m.lock);
        for (UINT32 i = 0; i < c->h->m.l.n; i++) kv_set(&x->headers, c->h->m.l.v[i].key, c->h->m.l.v[i].str, NULL, FALSE);
        ReleaseSRWLockShared(&c->h->m.lock);
        if (c->len && (x->body = zalloc(c->len))) { memcpy(x->body, c->data, c->len); x->body_len = c->len; }
        obj_release(&c->o);
    }
    if (!x->method || !x->url || !x->url[0]) {
        mfree(x->method); mfree(x->url); kv_clear(&x->headers); mfree(x->body); mfree(x);
        return E_INVALIDARG;
    }
    if (!(x->op = op_new(&IID_OpResponse, R_OBJ))) { mfree(x->method); mfree(x->url); kv_clear(&x->headers); mfree(x->body); mfree(x); return E_OUTOFMEMORY; }
    x->request = rq; obj_addref(&rq->o);
    x->op->cancel = send_cancel; x->op->cancel_ctx = x;
    obj_addref(&x->op->o);
    ret_itf(&x->op->o, out);   /* the caller's reference; x keeps the other */
    if (!QueueUserWorkItem(send_thread, x, WT_EXECUTELONGFUNCTION)) {
        HRESULT hr = last_error();
        x->op->cancel = NULL;
        op_finish(x->op, ST_ERROR, hr, NULL, NULL, 0);
        send_free(x);
    }
    return S_OK;
}

/* ---- HttpBaseProtocolFilter -------------------------------------------------------------------- */
struct filter { struct obj o; boolean8 redirects, allow_ui, decompress, use_proxy; UINT32 maxconn, maxversion, cookies; struct obj *cache; INT32 read_behavior, write_behavior; LONG validators; };
static void filter_destroy(struct obj *o) { obj_release(((struct filter *)o)->cache); }
static HRESULT WINAPI cc_get_read(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->read_behavior; return S_OK; }
static HRESULT WINAPI cc_put_read(struct itf *s, INT32 v) { ((struct filter *)s->obj)->read_behavior = v; return S_OK; }
static HRESULT WINAPI cc_get_write(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->write_behavior; return S_OK; }
static HRESULT WINAPI cc_put_write(struct itf *s, INT32 v) { ((struct filter *)s->obj)->write_behavior = v; return S_OK; }
static void *const cachecontrol_vtbl[] = { INSPECTABLE, (void *)cc_get_read, (void *)cc_put_read, (void *)cc_get_write, (void *)cc_put_write };
/* HttpCacheControl: a view of the filter's two settings (WinHTTP has no cache to apply them to) */
static HRESULT WINAPI f_get_CacheControl(struct itf *s, void **v)
{
    struct filter *f = (struct filter *)s->obj;
    if (!v) return E_POINTER;
    obj_addref(&f->o);
    *v = &f->o.ifs[7];
    return S_OK;
}
static HRESULT WINAPI f_get_redirects(struct itf *s, boolean8 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->redirects; return S_OK; }
static HRESULT WINAPI f_put_redirects(struct itf *s, boolean8 v) { ((struct filter *)s->obj)->redirects = v; return S_OK; }
static HRESULT WINAPI f_get_allow_ui(struct itf *s, boolean8 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->allow_ui; return S_OK; }
static HRESULT WINAPI f_put_allow_ui(struct itf *s, boolean8 v) { ((struct filter *)s->obj)->allow_ui = v; return S_OK; }
static HRESULT WINAPI f_get_decompress(struct itf *s, boolean8 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->decompress; return S_OK; }
static HRESULT WINAPI f_put_decompress(struct itf *s, boolean8 v) { ((struct filter *)s->obj)->decompress = v; return S_OK; }
static HRESULT WINAPI f_get_use_proxy(struct itf *s, boolean8 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->use_proxy; return S_OK; }
static HRESULT WINAPI f_put_use_proxy(struct itf *s, boolean8 v) { ((struct filter *)s->obj)->use_proxy = v; return S_OK; }
static HRESULT WINAPI f_get_maxconn(struct itf *s, UINT32 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->maxconn; return S_OK; }
static HRESULT WINAPI f_put_maxconn(struct itf *s, UINT32 v) { ((struct filter *)s->obj)->maxconn = v; return S_OK; }
static HRESULT WINAPI f_get_maxversion(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->maxversion; return S_OK; }
static HRESULT WINAPI f_put_maxversion(struct itf *s, INT32 v) { ((struct filter *)s->obj)->maxversion = v; return S_OK; }
static HRESULT WINAPI f_get_cookies(struct itf *s, INT32 *v) { if (!v) return E_POINTER; *v = ((struct filter *)s->obj)->cookies; return S_OK; }
static HRESULT WINAPI f_put_cookies(struct itf *s, INT32 v) { if (v != 0 && v != 1) return E_INVALIDARG; ((struct filter *)s->obj)->cookies = v; return S_OK; }   /* Default, NoCookies */
static HRESULT WINAPI f_obj_null(struct itf *s, void **v) { (void)s; if (!v) return E_POINTER; *v = NULL; return S_OK; }
/* ServerCustomValidationRequested: the handlers are never called (there is no server Certificate to hand
 * them), so a filter that has one fails its requests rather than skip the checks its owner asked for.
 * WinHTTP's own certificate validation applies either way; react-native-windows registers none. */
static HRESULT WINAPI f_add_validation(struct itf *s, void *handler, INT64 *token)
{
    if (!handler || !token) return E_POINTER;
    *token = InterlockedIncrement(&((struct filter *)s->obj)->validators);
    wh_log("ms365 webhttp: ServerCustomValidationRequested handler registered, the filter's requests will fail");
    return S_OK;
}
static HRESULT WINAPI f_remove_validation(struct itf *s, INT64 token)
{
    struct filter *f = (struct filter *)s->obj;
    (void)token;
    if (f->validators > 0) InterlockedDecrement(&f->validators);
    return S_OK;
}
static HRESULT WINAPI f_SendRequestAsync(struct itf *s, void *request, void **out)
{
    struct filter *f = (struct filter *)s->obj;
    if (f->validators > 0) {
        if (out) *out = NULL;
        notimpl("ServerCustomValidationRequested (custom server certificate validation)");
        return HRESULT_FROM_WIN32(ERROR_WINHTTP_SECURE_FAILURE);
    }
    return start_send(request, f->redirects, f->cookies != 1 /* HttpCookieUsageBehavior.NoCookies */, out);
}
NOTIMPL(f_get_CookieManager, "HttpBaseProtocolFilter.CookieManager")
NOTIMPL(f_get_IgnorableErrors, "HttpBaseProtocolFilter.IgnorableServerCertificateErrors")
static void *const filter_vtbl[] = { INSPECTABLE,
    (void *)f_get_redirects, (void *)f_put_redirects, (void *)f_get_allow_ui, (void *)f_put_allow_ui,
    (void *)f_get_decompress, (void *)f_put_decompress, (void *)f_get_CacheControl, (void *)f_get_CookieManager,
    (void *)f_obj_null, (void *)w_ok,                 /* ClientCertificate */
    (void *)f_get_IgnorableErrors, (void *)f_get_maxconn, (void *)f_put_maxconn,
    (void *)f_obj_null, (void *)w_ok,                 /* ProxyCredential */
    (void *)f_obj_null, (void *)w_ok,                 /* ServerCredential */
    (void *)f_get_use_proxy, (void *)f_put_use_proxy };
static void *const filter2_vtbl[] = { INSPECTABLE, (void *)f_get_maxversion, (void *)f_put_maxversion };
static void *const filter3_vtbl[] = { INSPECTABLE, (void *)f_get_cookies, (void *)f_put_cookies };
static void *const filter4_vtbl[] = { INSPECTABLE, (void *)f_add_validation, (void *)f_remove_validation, (void *)w_ok };   /* ..., ClearAuthenticationCache */
static void *const filter5_vtbl[] = { INSPECTABLE, (void *)f_obj_null };                                /* User */
static void *const httpfilter_vtbl[] = { INSPECTABLE, (void *)f_SendRequestAsync };
static struct obj *filter_new(void)
{
    struct filter *f = zalloc(sizeof(*f));
    if (!f) return NULL;
    void *const *vt[] = { filter_vtbl, filter2_vtbl, filter3_vtbl, filter4_vtbl, filter5_vtbl, httpfilter_vtbl, closable_vtbl, cachecontrol_vtbl };
    const GUID *ii[] = { &IID_Filter, &IID_Filter2, &IID_Filter3, &IID_Filter4, &IID_Filter5, &IID_HttpFilter, &IID_Closable, &IID_CacheControl };
    obj_init(&f->o, K_FILTER, L"Windows.Web.Http.Filters.HttpBaseProtocolFilter", filter_destroy, 8, vt, ii);
    f->redirects = f->decompress = f->use_proxy = 1; f->maxconn = 6; f->maxversion = 2 /* Http11 */;
    return &f->o;
}

/* ---- HttpClient -------------------------------------------------------------------------------- */
struct client { struct obj o; void *filter; struct headers *defaults; };
static void client_destroy(struct obj *o) { struct client *c = (struct client *)o; unk_release(c->filter); obj_release(c->defaults ? &c->defaults->m.o : NULL); }
static HRESULT client_send(struct client *c, void *request, void **out)
{
    struct request *rq = (struct request *)ours(request, K_REQUEST);
    if (!out) return E_POINTER;
    *out = NULL;
    if (!request) return E_INVALIDARG;
    if (rq) {   /* DefaultRequestHeaders fill in what the request does not set */
        AcquireSRWLockShared(&c->defaults->m.lock);
        AcquireSRWLockExclusive(&rq->h->m.lock);
        for (UINT32 i = 0; i < c->defaults->m.l.n; i++)
            if (kv_find(&rq->h->m.l, c->defaults->m.l.v[i].key) < 0) kv_set(&rq->h->m.l, c->defaults->m.l.v[i].key, c->defaults->m.l.v[i].str, NULL, FALSE);
        ReleaseSRWLockExclusive(&rq->h->m.lock);
        ReleaseSRWLockShared(&c->defaults->m.lock);
    }
    return VCALL(c->filter, 6, HRESULT (WINAPI *)(void *, void *, void **))(c->filter, request, out);   /* IHttpFilter::SendRequestAsync */
}
static HRESULT client_send_new(struct client *c, const WCHAR *method, void *uri, void *content, void **out)
{
    struct obj *m = method_new(method);
    struct request *rq = m ? request_new(itf0(m), uri) : NULL;
    obj_release(m);
    if (!rq) return E_OUTOFMEMORY;
    if (content) { rq->content = content; unk_addref(content); }
    HRESULT hr = client_send(c, itf0(&rq->o), out);
    obj_release(&rq->o);
    return hr;
}
static HRESULT WINAPI cl_DeleteAsync(struct itf *s, void *uri, void **out) { return client_send_new((struct client *)s->obj, L"DELETE", uri, NULL, out); }
static HRESULT WINAPI cl_GetAsync(struct itf *s, void *uri, void **out) { return client_send_new((struct client *)s->obj, L"GET", uri, NULL, out); }
static HRESULT WINAPI cl_GetAsync2(struct itf *s, void *uri, INT32 option, void **out) { (void)option; return cl_GetAsync(s, uri, out); }
static HRESULT WINAPI cl_PostAsync(struct itf *s, void *uri, void *content, void **out) { return client_send_new((struct client *)s->obj, L"POST", uri, content, out); }
static HRESULT WINAPI cl_PutAsync(struct itf *s, void *uri, void *content, void **out) { return client_send_new((struct client *)s->obj, L"PUT", uri, content, out); }
static HRESULT WINAPI cl_SendRequestAsync(struct itf *s, void *request, void **out) { return client_send((struct client *)s->obj, request, out); }
static HRESULT WINAPI cl_SendRequestAsync2(struct itf *s, void *request, INT32 option, void **out) { (void)option; return client_send((struct client *)s->obj, request, out); }
static HRESULT WINAPI cl_get_DefaultRequestHeaders(struct itf *s, void **out)
{
    struct client *c = (struct client *)s->obj;
    if (!out) return E_POINTER;
    obj_addref(&c->defaults->m.o); ret_itf(&c->defaults->m.o, out);
    return S_OK;
}
NOTIMPL(cl_GetBufferAsync, "HttpClient.GetBufferAsync")
NOTIMPL(cl_GetInputStreamAsync, "HttpClient.GetInputStreamAsync")
NOTIMPL(cl_GetStringAsync, "HttpClient.GetStringAsync")
NOTIMPL(cl_Try, "HttpClient.Try*Async")
static HRESULT WINAPI cl_get_PrivacyAnnotation(struct itf *s, HSTRING *v) { (void)s; if (!v) return E_POINTER; *v = NULL; return S_OK; }
static void *const client_vtbl[] = { INSPECTABLE, (void *)cl_DeleteAsync, (void *)cl_GetAsync, (void *)cl_GetAsync2,
    (void *)cl_GetBufferAsync, (void *)cl_GetInputStreamAsync, (void *)cl_GetStringAsync, (void *)cl_PostAsync, (void *)cl_PutAsync,
    (void *)cl_SendRequestAsync, (void *)cl_SendRequestAsync2, (void *)cl_get_DefaultRequestHeaders };
static void *const client2_vtbl[] = { INSPECTABLE, (void *)cl_Try, (void *)cl_Try, (void *)cl_Try, (void *)cl_Try, (void *)cl_Try,
    (void *)cl_Try, (void *)cl_Try, (void *)cl_Try, (void *)cl_Try, (void *)cl_Try };
static void *const client3_vtbl[] = { INSPECTABLE, (void *)cl_get_PrivacyAnnotation, (void *)w_ok };   /* DefaultPrivacyAnnotation */
/* takes a reference on filter (any IHttpFilter), or makes an HttpBaseProtocolFilter */
static struct obj *client_new(void *filter)
{
    struct client *c = zalloc(sizeof(*c));
    if (!c) return NULL;
    void *const *vt[] = { client_vtbl, client2_vtbl, client3_vtbl, closable_vtbl, stringable_vtbl };
    const GUID *ii[] = { &IID_Client, &IID_Client2, &IID_Client3, &IID_Closable, &IID_Stringable };
    obj_init(&c->o, K_CLIENT, L"Windows.Web.Http.HttpClient", client_destroy, 5, vt, ii);
    if (filter) { if (FAILED(unk_qi(filter, &IID_HttpFilter, &c->filter))) c->filter = NULL; }
    else { struct obj *f = filter_new(); c->filter = f ? &f->ifs[5] : NULL; }
    if (!c->filter || !(c->defaults = headers_new(H_REQUEST))) { obj_release(&c->o); return NULL; }
    return &c->o;
}

/* ---- activation factories (one immortal object per class) ----------------------------------- */
enum { C_FILTER, C_CLIENT, C_METHOD, C_REQUEST, C_RESPONSE, C_STRING, C_BUFFER, C_MEDIATYPE, C_READER, C_COUNT };
static const WCHAR *const class_names[C_COUNT] = {
    L"Windows.Web.Http.Filters.HttpBaseProtocolFilter", L"Windows.Web.Http.HttpClient", L"Windows.Web.Http.HttpMethod",
    L"Windows.Web.Http.HttpRequestMessage", L"Windows.Web.Http.HttpResponseMessage", L"Windows.Web.Http.HttpStringContent",
    L"Windows.Web.Http.HttpBufferContent", L"Windows.Web.Http.Headers.HttpMediaTypeHeaderValue", L"Windows.Storage.Streams.DataReader" };
struct factory { struct obj o; int cls; };
static HRESULT WINAPI fac_ActivateInstance(struct itf *s, void **out)
{
    struct obj *o = NULL;
    if (!out) return E_POINTER;
    *out = NULL;
    switch (((struct factory *)s->obj)->cls) {
    case C_FILTER: o = filter_new(); break;
    case C_CLIENT: o = client_new(NULL); break;
    case C_REQUEST: { struct request *r = request_new(NULL, NULL); o = r ? &r->o : NULL; break; }
    case C_RESPONSE: { struct response *r = response_new(200); o = r ? &r->o : NULL; break; }
    default: return notimpl("default constructor");
    }
    if (!o) return E_OUTOFMEMORY;
    ret_itf(o, out);
    return S_OK;
}
static HRESULT WINAPI fac_client_Create(struct itf *s, void *filter, void **out)
{
    (void)s;
    if (!out) return E_POINTER;
    struct obj *o = client_new(filter);
    if (!ret_itf(o, out)) return E_INVALIDARG;
    return S_OK;
}
static HRESULT method_out(const WCHAR *name, void **out)
{
    if (!out) return E_POINTER;
    return ret_itf(method_new(name), out) ? S_OK : E_OUTOFMEMORY;
}
static HRESULT WINAPI fac_method_Create(struct itf *s, HSTRING name, void **out)
{
    (void)s;
    WCHAR *n = from_hstring(name);
    HRESULT hr = !n ? E_OUTOFMEMORY : (is_token(n) && wcslen(n) < 24) ? method_out(n, out) : E_INVALIDARG;
    mfree(n);
    return hr;
}
static HRESULT WINAPI fac_delete(struct itf *s, void **o) { (void)s; return method_out(L"DELETE", o); }
static HRESULT WINAPI fac_get(struct itf *s, void **o) { (void)s; return method_out(L"GET", o); }
static HRESULT WINAPI fac_head(struct itf *s, void **o) { (void)s; return method_out(L"HEAD", o); }
static HRESULT WINAPI fac_options(struct itf *s, void **o) { (void)s; return method_out(L"OPTIONS", o); }
static HRESULT WINAPI fac_patch(struct itf *s, void **o) { (void)s; return method_out(L"PATCH", o); }
static HRESULT WINAPI fac_post(struct itf *s, void **o) { (void)s; return method_out(L"POST", o); }
static HRESULT WINAPI fac_put(struct itf *s, void **o) { (void)s; return method_out(L"PUT", o); }
static HRESULT WINAPI fac_request_Create(struct itf *s, void *method, void *uri, void **out)
{
    (void)s;
    if (!out) return E_POINTER;
    if (!method || !uri) { *out = NULL; return E_INVALIDARG; }
    struct request *r = request_new(method, uri);
    return ret_itf(r ? &r->o : NULL, out) ? S_OK : E_OUTOFMEMORY;
}
static HRESULT WINAPI fac_response_Create(struct itf *s, INT32 status, void **out)
{
    (void)s;
    if (!out) return E_POINTER;
    struct response *r = response_new(status);
    return ret_itf(r ? &r->o : NULL, out) ? S_OK : E_OUTOFMEMORY;
}
static HRESULT string_content(HSTRING text, INT32 encoding, const WCHAR *media, void **out)
{
    WCHAR *t = from_hstring(text);
    BYTE *data = NULL;
    UINT32 len = 0;
    if (!out) { mfree(t); return E_POINTER; }
    *out = NULL;
    if (!t) return E_OUTOFMEMORY;
    if (encoding == 0 /* Utf8 */) data = utf8_of(t, &len);
    else {
        len = (UINT32)(wcslen(t) * sizeof(WCHAR));
        if ((data = zalloc(len + 2))) {
            memcpy(data, t, len);
            if (encoding == 2 /* Utf16BE */) for (UINT32 i = 0; i + 1 < len; i += 2) { BYTE b = data[i]; data[i] = data[i + 1]; data[i + 1] = b; }
        }
    }
    mfree(t);
    if (!data) return E_OUTOFMEMORY;
    struct content *c = content_new(L"Windows.Web.Http.HttpStringContent", data, len);
    if (!c) return E_OUTOFMEMORY;
    WCHAR *ct = wconcat(media && media[0] ? media : L"text/plain", L"; charset=", encoding == 0 ? L"UTF-8" : encoding == 1 ? L"UTF-16LE" : L"UTF-16BE");
    if (ct) hdr_set(c->h, L"Content-Type", ct, FALSE);
    mfree(ct);
    ret_itf(&c->o, out);
    return S_OK;
}
static HRESULT WINAPI fac_string_Create(struct itf *s, HSTRING text, void **out) { (void)s; return string_content(text, 0, NULL, out); }
static HRESULT WINAPI fac_string_CreateEnc(struct itf *s, HSTRING text, INT32 enc, void **out) { (void)s; return string_content(text, enc, NULL, out); }
static HRESULT WINAPI fac_string_CreateEncMedia(struct itf *s, HSTRING text, INT32 enc, HSTRING media, void **out)
{
    (void)s;
    WCHAR *m = from_hstring(media);
    HRESULT hr = string_content(text, enc, m, out);
    mfree(m);
    return hr;
}
static HRESULT buffer_content(void *ibuffer, UINT32 offset, UINT32 count, BOOL whole, void **out)
{
    BYTE *src = NULL, *data;
    UINT32 len = 0;
    if (!out) return E_POINTER;
    *out = NULL;
    HRESULT hr = buffer_bytes(ibuffer, &src, &len, NULL);
    if (FAILED(hr)) return hr;
    if (whole) { offset = 0; count = len; }
    if (offset > len || count > len - offset) return E_INVALIDARG;
    if (!(data = zalloc(count ? count : 1))) return E_OUTOFMEMORY;
    if (count) memcpy(data, src + offset, count);
    struct content *c = content_new(L"Windows.Web.Http.HttpBufferContent", data, count);
    return ret_itf(c ? &c->o : NULL, out) ? S_OK : E_OUTOFMEMORY;
}
static HRESULT WINAPI fac_buffer_Create(struct itf *s, void *b, void **out) { (void)s; return buffer_content(b, 0, 0, TRUE, out); }
static HRESULT WINAPI fac_buffer_CreateRange(struct itf *s, void *b, UINT32 off, UINT32 n, void **out) { (void)s; return buffer_content(b, off, n, FALSE, out); }
static HRESULT WINAPI fac_mediatype_Create(struct itf *s, HSTRING media, void **out)
{
    (void)s;
    WCHAR *m = from_hstring(media);
    if (!out) { mfree(m); return E_POINTER; }
    HRESULT hr = (m && !wcschr(m, ';')) ? mediatype_parse(m, out) : E_INVALIDARG;
    mfree(m);
    return hr;
}
static HRESULT WINAPI fac_mediatype_Parse(struct itf *s, HSTRING input, void **out)
{
    (void)s;
    WCHAR *t = from_hstring(input);
    if (!out) { mfree(t); return E_POINTER; }
    HRESULT hr = t ? mediatype_parse(t, out) : E_OUTOFMEMORY;
    mfree(t);
    return hr;
}
static HRESULT WINAPI fac_mediatype_TryParse(struct itf *s, HSTRING input, void **out, boolean8 *ok)
{
    if (!out || !ok) return E_POINTER;
    HRESULT hr = fac_mediatype_Parse(s, input, out);
    *ok = SUCCEEDED(hr);
    return hr == E_INVALIDARG ? S_OK : hr;
}
static HRESULT WINAPI fac_reader_Create(struct itf *s, void *stream, void **out)
{
    (void)s;
    if (!out) return E_POINTER;
    if (!stream) { *out = NULL; return E_INVALIDARG; }
    struct reader *r = reader_new(stream);
    return ret_itf(r ? &r->o : NULL, out) ? S_OK : E_OUTOFMEMORY;
}
static HRESULT WINAPI fac_reader_FromBuffer(struct itf *s, void *ibuffer, void **out)
{
    (void)s;
    BYTE *data = NULL;
    UINT32 len = 0;
    if (!out) return E_POINTER;
    *out = NULL;
    HRESULT hr = buffer_bytes(ibuffer, &data, &len, NULL);
    if (FAILED(hr)) return hr;
    struct reader *r = reader_new(NULL);
    if (!r) return E_OUTOFMEMORY;
    if (len && !reader_push(r, data, len)) { obj_release(&r->o); return E_OUTOFMEMORY; }
    ret_itf(&r->o, out);
    return S_OK;
}
static void *const actfactory_vtbl[] = { INSPECTABLE, (void *)fac_ActivateInstance };
static void *const clientfactory_vtbl[] = { INSPECTABLE, (void *)fac_client_Create };
static void *const methodfactory_vtbl[] = { INSPECTABLE, (void *)fac_method_Create };
static void *const methodstatics_vtbl[] = { INSPECTABLE, (void *)fac_delete, (void *)fac_get, (void *)fac_head, (void *)fac_options, (void *)fac_patch, (void *)fac_post, (void *)fac_put };
static void *const requestfactory_vtbl[] = { INSPECTABLE, (void *)fac_request_Create };
static void *const responsefactory_vtbl[] = { INSPECTABLE, (void *)fac_response_Create };
static void *const stringfactory_vtbl[] = { INSPECTABLE, (void *)fac_string_Create, (void *)fac_string_CreateEnc, (void *)fac_string_CreateEncMedia };
static void *const bufferfactory_vtbl[] = { INSPECTABLE, (void *)fac_buffer_Create, (void *)fac_buffer_CreateRange };
static void *const mediatypefactory_vtbl[] = { INSPECTABLE, (void *)fac_mediatype_Create };
static void *const mediatypestatics_vtbl[] = { INSPECTABLE, (void *)fac_mediatype_Parse, (void *)fac_mediatype_TryParse };
static void *const readerfactory_vtbl[] = { INSPECTABLE, (void *)fac_reader_Create };
static void *const readerstatics_vtbl[] = { INSPECTABLE, (void *)fac_reader_FromBuffer };
static struct factory *g_factories[C_COUNT];

static struct factory *factory_new(int cls)
{
    struct factory *f = zalloc(sizeof(*f));
    if (!f) return NULL;
    void *const *vt[3] = { actfactory_vtbl }; const GUID *ii[3] = { &IID_ActivationFactory };
    int n = 1;
    switch (cls) {
    case C_CLIENT:    vt[n] = clientfactory_vtbl; ii[n++] = &IID_ClientFactory; break;
    case C_METHOD:    vt[n] = methodfactory_vtbl; ii[n++] = &IID_MethodFactory; vt[n] = methodstatics_vtbl; ii[n++] = &IID_MethodStatics; break;
    case C_REQUEST:   vt[n] = requestfactory_vtbl; ii[n++] = &IID_RequestFactory; break;
    case C_RESPONSE:  vt[n] = responsefactory_vtbl; ii[n++] = &IID_ResponseFactory; break;
    case C_STRING:    vt[n] = stringfactory_vtbl; ii[n++] = &IID_StringContentFactory; break;
    case C_BUFFER:    vt[n] = bufferfactory_vtbl; ii[n++] = &IID_BufferContentFactory; break;
    case C_MEDIATYPE: vt[n] = mediatypefactory_vtbl; ii[n++] = &IID_MediaTypeFactory; vt[n] = mediatypestatics_vtbl; ii[n++] = &IID_MediaTypeStatics; break;
    case C_READER:    vt[n] = readerfactory_vtbl; ii[n++] = &IID_DataReaderFactory; vt[n] = readerstatics_vtbl; ii[n++] = &IID_DataReaderStatics; break;
    }
    obj_init(&f->o, K_FACTORY, class_names[cls], NULL, n, vt, ii);
    f->o.immortal = TRUE; f->cls = cls;
    return f;
}

/* The activation factory for one of the classes above, QueryInterface'd to iid; REGDB_E_CLASSNOTREG
 * for any other class. */
HRESULT webhttp_get_factory(const WCHAR *name, REFIID iid, void **out)
{
    int cls;
    if (!out) return E_POINTER;
    *out = NULL;
    for (cls = 0; cls < C_COUNT; cls++) if (!wcscmp(name, class_names[cls])) break;
    if (cls == C_COUNT) return REGDB_E_CLASSNOTREG;
    if (!init()) return E_NOTIMPL;
    if (!g_factories[cls]) {
        struct factory *f = factory_new(cls);
        if (!f) return E_OUTOFMEMORY;
        if (InterlockedCompareExchangePointer((void **)&g_factories[cls], f, NULL) != NULL) mfree(f);
        else wh_log("ms365 webhttp: serving %ls", name);
    }
    return o_QueryInterface(&g_factories[cls]->o.ifs[0], iid, out);
}
