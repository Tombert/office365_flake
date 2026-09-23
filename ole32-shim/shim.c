/*
 * ole32.dll shim for Microsoft Office under Wine/Proton.
 *
 * 1. Exports: every export of the builtin ole32 is forwarded (see gen-def.py) to combase or to
 *    ole32_wine.dll (a renamed copy of the builtin), plus entry points the builtin lacks.
 *
 * 2. Missing-import patching: when a module imports a function that Wine's kernel32 (etc.) does
 *    not export at all, the Wine loader binds the import to a generated stub that raises a
 *    non-continuable exception on first call, which kills Office. Since ole32 is loaded early by
 *    every Office process, this DLL registers a loader notification and, for every module loaded
 *    from then on (and everything already loaded), rewrites the import-table slots of the
 *    functions listed in PATCHES to point at replacements that behave the way the real API does
 *    for an unprivileged caller.
 *
 * x86_64 only. Add entries to PATCHES as new "unimplemented function X.Y called" aborts turn up;
 * WINEDEBUG=+module (MS365_DEBUG=+module) lists every unresolved import as "No implementation for".
 */
#define COBJMACROS
#define INITGUID
#include <windows.h>
#include <msi.h>
typedef LPVOID HINTERNET;
#include <winternl.h>
#include <string.h>

/* ---- replacements ------------------------------------------------------------------------ */

/* mso30win32client GetProcAddress's this from ole32 and dereferences the result unconditionally.
 * Activation filters only matter for app-container activation; accept and ignore. */
__declspec(dllexport) HRESULT WINAPI CoRegisterActivationFilter(void *filter)
{
    (void)filter;
    return S_OK;
}

/* SetFileShortName needs SeRestorePrivilege on Windows; ordinary users get this failure and
 * Office carries on. mso.dll calls it while creating scratch files. */
static BOOL WINAPI my_SetFileShortNameW(HANDLE file, LPCWSTR name)
{
    (void)file; (void)name;
    SetLastError(ERROR_PRIVILEGE_NOT_HELD);
    return FALSE;
}
static BOOL WINAPI my_SetFileShortNameA(HANDLE file, LPCSTR name)
{
    (void)file; (void)name;
    SetLastError(ERROR_PRIVILEGE_NOT_HELD);
    return FALSE;
}

/* Windows 8+ packaged-app query; report "no packages" (ERROR_SUCCESS with zero results), which is
 * what an unpackaged desktop Office sees on a machine without the queried family installed. */
static LONG WINAPI my_FindPackagesByPackageFamily(PCWSTR family, UINT32 flags, UINT32 *count,
                                                  PWSTR *names, UINT32 *buflen, PWSTR buf, UINT32 *props)
{
    (void)family; (void)flags; (void)names; (void)buf; (void)props;
    if (count) *count = 0;
    if (buflen) *buflen = 0;
    return ERROR_SUCCESS;
}

/* Windows 8+ variant that also reports whether the timer was previously set; Wine has the plain
 * SetThreadpoolTimer, which is all aitrx.dll needs. */
static BOOL WINAPI my_SetThreadpoolTimerEx(PTP_TIMER timer, PFILETIME due, DWORD period, DWORD window)
{
    SetThreadpoolTimer(timer, due, period, window);
    return FALSE;
}

/* Wine's oleacc exports CreateStdAccessibleProxy{A,W} as aborting stubs. The proxy variant only
 * differs from CreateStdAccessibleObject (which Wine implements) by naming the window class, so
 * forward to that. Office's UI automation layer calls this once a document window is open. */
typedef HRESULT (WINAPI *pCreateStdAccessibleObject)(HWND, LONG, REFIID, void **);
static HRESULT WINAPI my_CreateStdAccessibleProxyW(HWND hwnd, LPCWSTR cls, LONG idObject, REFIID riid, void **ppv)
{
    (void)cls;
    HMODULE m = GetModuleHandleA("oleacc.dll");
    pCreateStdAccessibleObject fn = m ? (pCreateStdAccessibleObject)GetProcAddress(m, "CreateStdAccessibleObject") : NULL;
    if (!fn) { if (ppv) *ppv = NULL; return E_NOTIMPL; }
    return fn(hwnd, idObject, riid, ppv);
}
static HRESULT WINAPI my_CreateStdAccessibleProxyA(HWND hwnd, LPCSTR cls, LONG idObject, REFIID riid, void **ppv)
{
    (void)cls;
    return my_CreateStdAccessibleProxyW(hwnd, NULL, idObject, riid, ppv);
}

struct patch { const char *dll; const char *fn; void *repl; };

/* ---- replacements for functions Office delay-loads that Wine's DLLs do not export ----------
 * Delay-loads go through GetProcAddress at first call; a NULL result makes Office's delay-load
 * helper raise 0xC06D007F and the app dies. GetProcAddress is hooked (see my_GetProcAddress) to
 * fall back to these when the real lookup fails. */

/* user32: place a popup of `size` at `anchor`, kept inside the monitor's work area. */
static BOOL WINAPI my_CalculatePopupWindowPosition(const POINT *anchor, const SIZE *size, UINT flags, RECT *exclude, RECT *pos)
{
    (void)flags; (void)exclude;
    if (!anchor || !size || !pos) return FALSE;
    MONITORINFO mi; mi.cbSize = sizeof(mi);
    HMONITOR mon = MonitorFromPoint(*anchor, MONITOR_DEFAULTTONEAREST);
    RECT work = { 0, 0, 1920, 1080 };
    if (mon && GetMonitorInfoW(mon, &mi)) work = mi.rcWork;
    LONG x = anchor->x, y = anchor->y;
    if (x + size->cx > work.right) x = work.right - size->cx;
    if (y + size->cy > work.bottom) y = work.bottom - size->cy;
    if (x < work.left) x = work.left;
    if (y < work.top) y = work.top;
    pos->left = x; pos->top = y; pos->right = x + size->cx; pos->bottom = y + size->cy;
    return TRUE;
}
static BOOL WINAPI my_InheritWindowMonitor(HWND hwnd, HWND inherit) { (void)hwnd; (void)inherit; return TRUE; }

/* slc.dll: forward to our sppc.dll */
static HRESULT WINAPI my_SLGetGenuineInformation(const void *id, LPCWSTR name, int *type, UINT *size, BYTE **data)
{
    typedef HRESULT (WINAPI *fn_t)(const void *, LPCWSTR, int *, UINT *, BYTE **);
    HMODULE m = LoadLibraryA("sppc.dll");
    fn_t fn = m ? (fn_t)GetProcAddress(m, "SLGetGenuineInformation") : NULL;
    if (!fn) { if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return (HRESULT)0xC004F012; }
    return fn(id, name, type, size, data);
}

/* misc: report "not supported" the way the real API does when the feature is absent */
static HRESULT WINAPI my_DeleteAppContainerProfile(LPCWSTR name) { (void)name; return S_OK; }
static DWORD WINAPI my_DavFlushFile(HANDLE h) { (void)h; return ERROR_NOT_SUPPORTED; }
static DWORD WINAPI my_DavGetExtendedError(HANDLE h, DWORD *err, LPWSTR buf, DWORD *len) { (void)h; if (err) *err = 0; if (buf && len && *len) buf[0] = 0; if (len) *len = 0; return ERROR_NOT_SUPPORTED; }
static BOOL WINAPI my_CertSelectCertificateChains(void *a, void *b, void *c, void *d, void *e, DWORD *count, void **chains)
{ (void)a; (void)b; (void)c; (void)d; (void)e; if (count) *count = 0; if (chains) *chains = NULL; SetLastError(ERROR_NOT_SUPPORTED); return FALSE; }
static void WINAPI my_CertFreeCertificateChainList(void *chains) { (void)chains; }
static BOOL WINAPI my_CryptRetrieveTimeStamp(void) { SetLastError(ERROR_NOT_SUPPORTED); return FALSE; }
static BOOL WINAPI my_CryptVerifyTimeStampSignature(void) { SetLastError(ERROR_NOT_SUPPORTED); return FALSE; }
static HRESULT WINAPI my_CertSelectionGetSerializedBlob(void) { return E_NOTIMPL; }

static const struct patch DELAYPATCHES[] = {
    { "user32.dll",   "CalculatePopupWindowPosition",  (void *)my_CalculatePopupWindowPosition },
    { "user32.dll",   "InheritWindowMonitor",          (void *)my_InheritWindowMonitor },
    { "slc.dll",      "SLGetGenuineInformation",       (void *)my_SLGetGenuineInformation },
    { "userenv.dll",  "DeleteAppContainerProfile",     (void *)my_DeleteAppContainerProfile },
    { "netapi32.dll", "DavFlushFile",                  (void *)my_DavFlushFile },
    { "netapi32.dll", "DavGetExtendedError",           (void *)my_DavGetExtendedError },
    { "crypt32.dll",  "CertSelectCertificateChains",   (void *)my_CertSelectCertificateChains },
    { "crypt32.dll",  "CertFreeCertificateChainList",  (void *)my_CertFreeCertificateChainList },
    { "crypt32.dll",  "CryptRetrieveTimeStamp",        (void *)my_CryptRetrieveTimeStamp },
    { "crypt32.dll",  "CryptVerifyTimeStampSignature", (void *)my_CryptVerifyTimeStampSignature },
    { "cryptui.dll",  "CertSelectionGetSerializedBlob",(void *)my_CertSelectionGetSerializedBlob },
};
#define NDELAYPATCHES (sizeof(DELAYPATCHES) / sizeof(DELAYPATCHES[0]))

/* ---- kernel32: GetDllDirectory when no directory is set ----------------------------------
 * GetDllDirectory returns 0 both for "no DLL directory set" and for failure; callers tell the two
 * apart with GetLastError, which Windows leaves at ERROR_SUCCESS in the first case. Wine's returns 0
 * and does not touch the last error, so whatever the thread last failed at (Office's C2R layer
 * leaves ERROR_ENVVAR_NOT_FOUND lying around) reads as a failure. Excel's Solver checks exactly
 * this, through VBA's Err.LastDllError, before loading SOLVER32.DLL and gives up with "Solver
 * encountered an error value in the Objective Cell" on the first Solve of a session. VBA reaches
 * the function through GetProcAddress, so my_GetProcAddress hands out the wrapper as well. */
typedef DWORD (WINAPI *pGetDllDirectoryA)(DWORD, LPSTR);
typedef DWORD (WINAPI *pGetDllDirectoryW)(DWORD, LPWSTR);
static pGetDllDirectoryA real_GetDllDirectoryA;
static pGetDllDirectoryW real_GetDllDirectoryW;
static DWORD WINAPI my_GetDllDirectoryA(DWORD len, LPSTR buf)
{
    DWORD r = real_GetDllDirectoryA ? real_GetDllDirectoryA(len, buf) : 0;
    if (!r) SetLastError(ERROR_SUCCESS);   /* 0: an empty directory (or, only in theory, out of memory) */
    return r;
}
static DWORD WINAPI my_GetDllDirectoryW(DWORD len, LPWSTR buf)
{
    DWORD r = real_GetDllDirectoryW ? real_GetDllDirectoryW(len, buf) : 0;
    if (!r) SetLastError(ERROR_SUCCESS);
    return r;
}

typedef FARPROC (WINAPI *pGetProcAddress)(HMODULE, LPCSTR);
static pGetProcAddress real_GetProcAddress;
typedef DWORD (WINAPI *pGetModuleBaseNameA_t)(HANDLE, HMODULE, LPSTR, DWORD);
static pGetModuleBaseNameA_t real_GetModuleBaseNameA;

static const struct patch *wrapped_lookup(HMODULE h, LPCSTR name);
static const struct patch *wrapped_lookup_ordinal(HMODULE h, WORD ordinal);
/* MS365_TRACE_MODULE=<dll>: log every export lookup into that module and every failed lookup
 * anywhere (diagnostics for probes such as Office loading a proofing engine and dropping it). */
static HMODULE g_trace_mod; static char g_trace_mod_name[64]; static LONG g_gpa_logs;
static void gpa_log(HMODULE h, LPCSTR name, FARPROC p)
{
    char buf[300], mod[128] = "?";
    if (InterlockedIncrement(&g_gpa_logs) > 400) return;
    if (real_GetModuleBaseNameA) real_GetModuleBaseNameA(GetCurrentProcess(), h, mod, sizeof(mod));
    if (((ULONG_PTR)name >> 16) != 0) wsprintfA(buf, "ms365 ole32 shim: GetProcAddress %s!%s -> %p", mod, name, p);
    else wsprintfA(buf, "ms365 ole32 shim: GetProcAddress %s!#%u -> %p", mod, (UINT)(ULONG_PTR)name, p);
    OutputDebugStringA(buf);
}
static FARPROC WINAPI my_GetProcAddress(HMODULE h, LPCSTR name)
{
    FARPROC p = real_GetProcAddress(h, name);
    if (g_trace_mod_name[0] && name && (h == g_trace_mod || !p)) gpa_log(h, name, p);
    /* by address rather than by module, so a forwarder or an api-set alias resolves the same */
    if (p && p == (FARPROC)real_GetDllDirectoryA) return (FARPROC)my_GetDllDirectoryA;
    if (p && p == (FARPROC)real_GetDllDirectoryW) return (FARPROC)my_GetDllDirectoryW;
    if (name && ((ULONG_PTR)name >> 16) != 0) {
        /* functions we wrap (see PATCHES) must be wrapped for dynamic lookups too */
        const struct patch *w = p ? wrapped_lookup(h, name) : NULL;
        if (w) return (FARPROC)w->repl;
    } else if (name && p) {
        /* ordinal lookups into wrapped modules (Office links msi.dll by ordinal) */
        const struct patch *w = wrapped_lookup_ordinal(h, (WORD)(ULONG_PTR)name);
        if (w) return (FARPROC)w->repl;
    }
    if (p || !name || ((ULONG_PTR)name >> 16) == 0) return p;   /* ordinal lookups pass through */
    char mod[128] = "";
    if (real_GetModuleBaseNameA) real_GetModuleBaseNameA(GetCurrentProcess(), h, mod, sizeof(mod));
    for (size_t i = 0; i < NDELAYPATCHES; i++) {
        if (lstrcmpiA(mod, DELAYPATCHES[i].dll) == 0 && strcmp(name, DELAYPATCHES[i].fn) == 0) {
            char buf[256]; wsprintfA(buf, "ms365 ole32 shim: GetProcAddress fallback %s!%s", mod, name); OutputDebugStringA(buf);
            return (FARPROC)DELAYPATCHES[i].repl;
        }
    }
    return NULL;
}

/* Exported stubs to overwrite in place (12-byte "mov rax, imm64; jmp rax") when their DLL loads.
 * Unlike PATCHES these exist in the export table, but calling them aborts the process. */
static const struct patch STUBPATCHES[] = {
    { "oleacc.dll", "CreateStdAccessibleProxyW", (void *)my_CreateStdAccessibleProxyW },
    { "oleacc.dll", "CreateStdAccessibleProxyA", (void *)my_CreateStdAccessibleProxyA },
};
#define NSTUBPATCHES (sizeof(STUBPATCHES) / sizeof(STUBPATCHES[0]))

/* ---- crypt32: the first chain on a default engine ----------------------------------------------
 * Wine's crypt32 builds the default chain engines (HCCE_CURRENT_USER / HCCE_LOCAL_MACHINE) on first
 * use. When two threads verify a certificate at the same moment (OneNote's first HTTPS requests at
 * start), both build one and the loser frees its copy at once; closing that copy's registry store
 * (RegCloseKey, then CloseHandle on the key's change-notification event) crashes inside Office's
 * App-V layer (AppVIsvSubsystems64 null write) in about every second OneNote start. Serialise
 * CertGetCertificateChain on a default engine until that engine exists, so there is no loser. */
typedef BOOL (WINAPI *pCertGetCertificateChain)(void *, const void *, void *, void *, void *, DWORD, void *, void **);
static pCertGetCertificateChain real_CertGetCertificateChain;
static SRWLOCK g_chain_lock = SRWLOCK_INIT;
static LONG g_chain_ready;      /* bit 0: current-user engine built, bit 1: local-machine engine */
static BOOL WINAPI my_CertGetCertificateChain(void *engine, const void *cert, void *time, void *store,
                                              void *para, DWORD flags, void *reserved, void **chain)
{
    pCertGetCertificateChain fn = real_CertGetCertificateChain;
    LONG bit = !engine ? 1 : engine == (void *)1 ? 2 : 0;     /* HCCE_CURRENT_USER, HCCE_LOCAL_MACHINE */
    BOOL ret;
    if (!fn) { SetLastError(ERROR_PROC_NOT_FOUND); return FALSE; }
    if (!bit || (g_chain_ready & bit)) return fn(engine, cert, time, store, para, flags, reserved, chain);
    AcquireSRWLockExclusive(&g_chain_lock);
    ret = fn(engine, cert, time, store, para, flags, reserved, chain);
    InterlockedOr(&g_chain_ready, bit);
    ReleaseSRWLockExclusive(&g_chain_lock);
    return ret;
}

/* ---- virtdisk: GetStorageDependencyInformation ----------------------------------------------------
 * Before Office opens a macro file (an .xlam add-in such as Solver, .xlsm, .docm) it asks whether the
 * file's volume is backed by a mounted ISO/VHD, where macros are blocked: it opens the volume, calls
 * GetStorageDependencyInformation(GET_STORAGE_DEPENDENCY_FLAG_HOST_VOLUMES), retries on
 * ERROR_INSUFFICIENT_BUFFER, treats any other error as "no backing image", and on success reads entry 0's
 * host volume and relative path without looking at NumberEntries. Wine's stub returns success with
 * NumberEntries = 0 and the buffer untouched, so mso copied strings through garbage pointers (Excel crashed
 * while enabling Solver). Nothing in a Wine prefix sits on a virtual disk: give Windows' answer for an
 * ordinary volume. */
#define MY_ERROR_VIRTDISK_NOT_VIRTUAL_DISK 0xC03A0015
static DWORD WINAPI my_GetStorageDependencyInformation(HANDLE obj, int flags, ULONG size, void *info, ULONG *used)
{
    (void)obj; (void)flags; (void)size; (void)info;
    if (used) *used = 0;
    return MY_ERROR_VIRTDISK_NOT_VIRTUAL_DISK;
}

/* ---- WinHTTP options Wine has not implemented -------------------------------------------
 * Wine's winhttp fails WinHttpSetOption/WinHttpQueryOption with ERROR_WINHTTP_INVALID_OPTION
 * (12009) for options it doesn't know, and Office's OneAuth treats that as a failed request
 * (sign-in dies right after the password step with code 12009). The options it uses are tuning
 * knobs: IPv6 fast fallback (140) and the autologon policy query (77), for which the default
 * security level is the right answer. */
#define MY_WINHTTP_OPTION_AUTOLOGON_POLICY     77
#define MY_WINHTTP_OPTION_IPV6_FAST_FALLBACK   140
#define MY_ERROR_WINHTTP_INVALID_OPTION        12009
/* Wine reports an unknown option as 12009 (session/connect handles) or ERROR_INVALID_PARAMETER 87
 * (request handles); ERROR_WINHTTP_INCORRECT_HANDLE_TYPE 12018 for handles with no query table. */
static int unknown_option_error(DWORD err) { return err == MY_ERROR_WINHTTP_INVALID_OPTION || err == ERROR_INVALID_PARAMETER || err == 12018; }
typedef BOOL (WINAPI *pWinHttpSetOption)(HANDLE, DWORD, LPVOID, DWORD);
typedef BOOL (WINAPI *pWinHttpQueryOption)(HANDLE, DWORD, LPVOID, LPDWORD);
static pWinHttpSetOption real_WinHttpSetOption;
static pWinHttpQueryOption real_WinHttpQueryOption;
static BOOL WINAPI my_WinHttpSetOption(HANDLE h, DWORD option, LPVOID buf, DWORD len)
{
    pWinHttpSetOption fn = real_WinHttpSetOption;
    if (fn && fn(h, option, buf, len)) return TRUE;
    DWORD err = GetLastError();
    if (unknown_option_error(err) && (option >= 128 || option == MY_WINHTTP_OPTION_AUTOLOGON_POLICY)) {
        char b[96]; wsprintfA(b, "ms365 ole32 shim: WinHttpSetOption %lu accepted (unimplemented in Wine)", option); OutputDebugStringA(b);
        SetLastError(0); return TRUE; /* newer tuning options and the autologon policy: accept */
    }
    SetLastError(err);
    return FALSE;
}
static BOOL WINAPI my_WinHttpQueryOption(HANDLE h, DWORD option, LPVOID buf, LPDWORD len)
{
    pWinHttpQueryOption fn = real_WinHttpQueryOption;
    if (fn && fn(h, option, buf, len)) return TRUE;
    DWORD err = GetLastError();
    if (unknown_option_error(err) && option == MY_WINHTTP_OPTION_AUTOLOGON_POLICY && len) {
        if (!buf || *len < sizeof(DWORD)) { *len = sizeof(DWORD); SetLastError(ERROR_INSUFFICIENT_BUFFER); return FALSE; }
        *(DWORD *)buf = 1; /* WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM, the Windows default */
        OutputDebugStringA("ms365 ole32 shim: WinHttpQueryOption 77 answered with the default autologon policy");
        *len = sizeof(DWORD);
        SetLastError(0);
        return TRUE;
    }
    SetLastError(err);
    return FALSE;
}

/* ---- WinINet options Wine has not implemented -------------------------------------------
 * Office's Microsoft-account ticket requests (the legacy MSA path, used once WAM and OneAuth are
 * off) go through wininet. Wine's InternetSetOption rejects options it has not implemented with
 * ERROR_INTERNET_INVALID_OPTION (12009, the same number as its winhttp cousin) and the request is
 * abandoned with sign-in error tag 53u4r. The options seen are tuning knobs (11 = listen timeout),
 * so unknown ones are accepted. Query failures are only traced. */
typedef BOOL (WINAPI *pInternetSetOption)(HINTERNET, DWORD, LPVOID, DWORD);
typedef BOOL (WINAPI *pInternetQueryOption)(HINTERNET, DWORD, LPVOID, LPDWORD);
static pInternetSetOption real_InternetSetOptionW, real_InternetSetOptionA;
static pInternetQueryOption real_InternetQueryOptionW, real_InternetQueryOptionA;
static BOOL inet_set_option(pInternetSetOption fn, const char *who, HINTERNET h, DWORD option, LPVOID buf, DWORD len)
{
    if (fn && fn(h, option, buf, len)) return TRUE;
    DWORD err = GetLastError();
    if (err == MY_ERROR_WINHTTP_INVALID_OPTION) {
        char b[96]; wsprintfA(b, "ms365 ole32 shim: %s %lu accepted (unimplemented in Wine)", who, option); OutputDebugStringA(b);
        SetLastError(0); return TRUE;
    }
    SetLastError(err);
    return FALSE;
}
static BOOL inet_query_option(pInternetQueryOption fn, const char *who, HINTERNET h, DWORD option, LPVOID buf, LPDWORD len)
{
    if (fn && fn(h, option, buf, len)) return TRUE;
    DWORD err = GetLastError();
    if (err == MY_ERROR_WINHTTP_INVALID_OPTION) {
        char b[96]; wsprintfA(b, "ms365 ole32 shim: %s %lu failed (unimplemented in Wine)", who, option); OutputDebugStringA(b);
    }
    SetLastError(err);
    return FALSE;
}
static BOOL WINAPI my_InternetSetOptionW(HINTERNET h, DWORD o, LPVOID b, DWORD l)     { return inet_set_option(real_InternetSetOptionW, "InternetSetOptionW", h, o, b, l); }
static BOOL WINAPI my_InternetSetOptionA(HINTERNET h, DWORD o, LPVOID b, DWORD l)     { return inet_set_option(real_InternetSetOptionA, "InternetSetOptionA", h, o, b, l); }
static BOOL WINAPI my_InternetQueryOptionW(HINTERNET h, DWORD o, LPVOID b, LPDWORD l) { return inet_query_option(real_InternetQueryOptionW, "InternetQueryOptionW", h, o, b, l); }
static BOOL WINAPI my_InternetQueryOptionA(HINTERNET h, DWORD o, LPVOID b, LPDWORD l) { return inet_query_option(real_InternetQueryOptionA, "InternetQueryOptionA", h, o, b, l); }


/* ---- WinRT statics Wine lacks --------------------------------------------------------------
 * Office's licensing dialogs (React Native windows in osf99 / react-native-win32) ask the
 * Windows.Globalization.Language factory for ILanguageStatics (IsWellFormed,
 * CurrentInputMethodLanguageTag). Wine's factory only implements ILanguageFactory and answers
 * E_NOINTERFACE, which the dialog reports as a fatal error and Word exits. RoGetActivationFactory
 * is wrapped (static imports and the delay-load lookups through api-ms-win-core-winrt-l1-1-0)
 * and that one request is served by a tiny statics object. */
typedef void *HSTRING;
typedef HRESULT (WINAPI *pRoGetActivationFactory)(HSTRING, REFIID, void **);
typedef const WCHAR *(WINAPI *pWindowsGetStringRawBuffer)(HSTRING, UINT32 *);
typedef HRESULT (WINAPI *pWindowsCreateString)(const WCHAR *, UINT32, HSTRING *);
static pRoGetActivationFactory real_RoGetActivationFactory;
static pWindowsGetStringRawBuffer p_WindowsGetStringRawBuffer;
static pWindowsCreateString p_WindowsCreateString;
static const GUID MY_IID_IInspectable    = { 0xaf86e2e0, 0xb12d, 0x4c6a, { 0x9c, 0x5a, 0xd7, 0xaa, 0x65, 0x10, 0x1e, 0x90 } };
static const GUID MY_IID_ILanguageStatics = { 0xb23cd557, 0x0865, 0x46d4, { 0x89, 0xb8, 0xd5, 0x9b, 0xe8, 0x99, 0x0f, 0x0d } };

typedef struct lang_statics { const struct lang_statics_vtbl *vtbl; } lang_statics;
struct lang_statics_vtbl {
    HRESULT (WINAPI *QueryInterface)(lang_statics *, REFIID, void **);
    ULONG   (WINAPI *AddRef)(lang_statics *);
    ULONG   (WINAPI *Release)(lang_statics *);
    HRESULT (WINAPI *GetIids)(lang_statics *, ULONG *, GUID **);
    HRESULT (WINAPI *GetRuntimeClassName)(lang_statics *, HSTRING *);
    HRESULT (WINAPI *GetTrustLevel)(lang_statics *, int *);
    HRESULT (WINAPI *IsWellFormed)(lang_statics *, HSTRING, BOOLEAN *);
    HRESULT (WINAPI *get_CurrentInputMethodLanguageTag)(lang_statics *, HSTRING *);
};
static HRESULT WINAPI ls_QueryInterface(lang_statics *this, REFIID iid, void **out)
{
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &IID_IUnknown) || IsEqualGUID(iid, &MY_IID_IInspectable) || IsEqualGUID(iid, &MY_IID_ILanguageStatics)) { *out = this; return S_OK; }
    *out = NULL; return E_NOINTERFACE;
}
static ULONG WINAPI ls_AddRef(lang_statics *this) { (void)this; return 2; }
static ULONG WINAPI ls_Release(lang_statics *this) { (void)this; return 1; }
static HRESULT WINAPI ls_GetIids(lang_statics *this, ULONG *n, GUID **iids) { (void)this; if (n) *n = 0; if (iids) *iids = NULL; return S_OK; }
static HRESULT WINAPI ls_GetRuntimeClassName(lang_statics *this, HSTRING *name)
{ (void)this; static const WCHAR cls[] = L"Windows.Globalization.Language"; return p_WindowsCreateString ? p_WindowsCreateString(cls, 30, name) : E_NOTIMPL; }
static HRESULT WINAPI ls_GetTrustLevel(lang_statics *this, int *level) { (void)this; if (level) *level = 0; return S_OK; }
static HRESULT WINAPI ls_IsWellFormed(lang_statics *this, HSTRING tag, BOOLEAN *result)
{
    (void)this;
    UINT32 len = 0; const WCHAR *t = (tag && p_WindowsGetStringRawBuffer) ? p_WindowsGetStringRawBuffer(tag, &len) : NULL;
    if (result) *result = (t && len > 0 && len < 64);
    return S_OK;
}
static HRESULT WINAPI ls_get_CurrentInputMethodLanguageTag(lang_statics *this, HSTRING *value)
{
    (void)this;
    WCHAR tag[LOCALE_NAME_MAX_LENGTH] = L"en-US";
    if (!GetUserDefaultLocaleName(tag, LOCALE_NAME_MAX_LENGTH) || !tag[0]) lstrcpyW(tag, L"en-US");
    return p_WindowsCreateString ? p_WindowsCreateString(tag, lstrlenW(tag), value) : E_NOTIMPL;
}
static const struct lang_statics_vtbl lang_statics_vtbl = {
    ls_QueryInterface, ls_AddRef, ls_Release, ls_GetIids, ls_GetRuntimeClassName, ls_GetTrustLevel,
    ls_IsWellFormed, ls_get_CurrentInputMethodLanguageTag,
};
static lang_statics lang_statics_obj = { &lang_statics_vtbl };

static void resolve_winrt(void)
{
    if (real_RoGetActivationFactory) return;
    HMODULE cb = GetModuleHandleA("combase.dll");
    if (!cb) cb = LoadLibraryA("combase.dll");
    if (!cb) return;
    p_WindowsGetStringRawBuffer = (pWindowsGetStringRawBuffer)real_GetProcAddress(cb, "WindowsGetStringRawBuffer");
    p_WindowsCreateString = (pWindowsCreateString)real_GetProcAddress(cb, "WindowsCreateString");
    real_RoGetActivationFactory = (pRoGetActivationFactory)real_GetProcAddress(cb, "RoGetActivationFactory");
}

/* Windows.Data.Json.JsonObject statics (Parse/TryParse). Wine's windows.web has JsonValue statics
 * and JsonObject instances but no IJsonObjectStatics; Office's licensing code parses the licence
 * model with JsonObject.Parse. Built on top of Wine's JsonValue.Parse + IJsonValue.GetObject. */
static const GUID MY_IID_IJsonObjectStatics = { 0x2289f159, 0x54de, 0x45d8, { 0xab, 0xcc, 0x22, 0x60, 0x3f, 0xa0, 0x66, 0xa0 } };
static const GUID MY_IID_IJsonValueStatics  = { 0x5f6b544a, 0x2f53, 0x48e1, { 0x91, 0xa3, 0xf7, 0x8b, 0x50, 0xa6, 0x34, 0x5c } };
typedef struct json_statics { const struct json_statics_vtbl *vtbl; } json_statics;
struct json_statics_vtbl {
    HRESULT (WINAPI *QueryInterface)(json_statics *, REFIID, void **);
    ULONG   (WINAPI *AddRef)(json_statics *);
    ULONG   (WINAPI *Release)(json_statics *);
    HRESULT (WINAPI *GetIids)(json_statics *, ULONG *, GUID **);
    HRESULT (WINAPI *GetRuntimeClassName)(json_statics *, HSTRING *);
    HRESULT (WINAPI *GetTrustLevel)(json_statics *, int *);
    HRESULT (WINAPI *Parse)(json_statics *, HSTRING, void **);
    HRESULT (WINAPI *TryParse)(json_statics *, HSTRING, void **, BOOLEAN *);
};
/* vtable slots of the Wine-provided interfaces we call through */
typedef struct { void **vtbl; } rt_obj;
#define RT_CALL(obj, idx) ((obj)->vtbl[idx])
typedef HRESULT (WINAPI *fn_rt_qi)(void *, REFIID, void **);
typedef ULONG   (WINAPI *fn_rt_release)(void *);
typedef HRESULT (WINAPI *fn_jvs_parse)(void *, HSTRING, void **);          /* IJsonValueStatics slot 6 */
typedef HRESULT (WINAPI *fn_jv_getobject)(void *, void **);               /* IJsonValue slot 12 */
static HRESULT json_object_parse(HSTRING input, void **out)
{
    resolve_winrt();
    if (!real_RoGetActivationFactory || !p_WindowsCreateString) return E_NOTIMPL;
    HSTRING cls = NULL; void *statics = NULL; void *value = NULL; HRESULT hr;
    static const WCHAR name[] = L"Windows.Data.Json.JsonValue";
    if (FAILED(hr = p_WindowsCreateString(name, 27, &cls))) return hr;
    hr = real_RoGetActivationFactory(cls, &MY_IID_IJsonValueStatics, &statics);
    if (SUCCEEDED(hr)) {
        hr = ((fn_jvs_parse)RT_CALL((rt_obj *)statics, 6))(statics, input, &value);
        if (SUCCEEDED(hr)) {
            hr = ((fn_jv_getobject)RT_CALL((rt_obj *)value, 12))(value, out);
            ((fn_rt_release)RT_CALL((rt_obj *)value, 2))(value);
        }
        ((fn_rt_release)RT_CALL((rt_obj *)statics, 2))(statics);
    }
    return hr;
}
static HRESULT WINAPI js_QueryInterface(json_statics *this, REFIID iid, void **out)
{
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &IID_IUnknown) || IsEqualGUID(iid, &MY_IID_IInspectable) || IsEqualGUID(iid, &MY_IID_IJsonObjectStatics)) { *out = this; return S_OK; }
    *out = NULL; return E_NOINTERFACE;
}
static ULONG WINAPI js_AddRef(json_statics *this) { (void)this; return 2; }
static ULONG WINAPI js_Release(json_statics *this) { (void)this; return 1; }
static HRESULT WINAPI js_GetIids(json_statics *this, ULONG *n, GUID **iids) { (void)this; if (n) *n = 0; if (iids) *iids = NULL; return S_OK; }
static HRESULT WINAPI js_GetRuntimeClassName(json_statics *this, HSTRING *name)
{ (void)this; static const WCHAR cls[] = L"Windows.Data.Json.JsonObject"; return p_WindowsCreateString ? p_WindowsCreateString(cls, 28, name) : E_NOTIMPL; }
static HRESULT WINAPI js_GetTrustLevel(json_statics *this, int *level) { (void)this; if (level) *level = 0; return S_OK; }
static HRESULT WINAPI js_Parse(json_statics *this, HSTRING input, void **out)
{
    (void)this;
    if (!out) return E_POINTER;
    *out = NULL;
    HRESULT hr = json_object_parse(input, out);
    if (FAILED(hr)) { char b[96]; wsprintfA(b, "ms365 ole32 shim: JsonObject.Parse failed 0x%08lx", hr); OutputDebugStringA(b); }
    return hr;
}
static HRESULT WINAPI js_TryParse(json_statics *this, HSTRING input, void **out, BOOLEAN *ok)
{
    (void)this;
    if (!out || !ok) return E_POINTER;
    *out = NULL;
    *ok = SUCCEEDED(json_object_parse(input, out));
    return S_OK;
}
static const struct json_statics_vtbl json_statics_vtbl = {
    js_QueryInterface, js_AddRef, js_Release, js_GetIids, js_GetRuntimeClassName, js_GetTrustLevel, js_Parse, js_TryParse,
};
static json_statics json_statics_obj = { &json_statics_vtbl };

/* ---- Windows.Web.Http stand-ins ----------------------------------------------------------------
 * OneNote's React Native UI (react-native-win32) creates a Windows.Web.Http.Filters.
 * HttpBaseProtocolFilter and an HttpClient when its networking module starts. Wine has no
 * Windows.Web.Http at all; the failed activation throws on the JavaScript thread, Office's error
 * reporting then loops there, and every React Native surface (the notebook navigation pane) stays
 * blank. These stand-ins construct like the real classes (filter settings, cache control, the
 * certificate-validation event, HttpMethod), and every request fails with "cannot connect"
 * (WININET_E_CANNOT_CONNECT), which React Native reports as an ordinary network error. Layouts
 * follow Windows.Foundation.UniversalApiContract.winmd; every out parameter gets its exact width
 * (a WinRT boolean is one byte). */
#define HTTP_E_CANNOT_CONNECT ((HRESULT)0x80072EFD)
static const GUID MY_IID_IActivationFactory      = { 0x00000035, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
static const GUID MY_IID_IClosable               = { 0x30d5a829, 0x7fa4, 0x4026, { 0x83, 0xbb, 0xd7, 0x5b, 0xae, 0x4e, 0xa9, 0x9e } };
static const GUID MY_IID_IStringable             = { 0x96369f54, 0x8eb6, 0x48f0, { 0xab, 0xce, 0xc1, 0xb2, 0x11, 0xe6, 0x27, 0xc3 } };
static const GUID MY_IID_IHttpBaseProtocolFilter = { 0x71c89b09, 0xe131, 0x4b54, { 0xa5, 0x3c, 0xeb, 0x43, 0xff, 0x37, 0xe9, 0xbb } };
static const GUID MY_IID_IHttpBaseProtocolFilter2= { 0x2ec30013, 0x9427, 0x4900, { 0xa0, 0x17, 0xfa, 0x7d, 0xa3, 0xb5, 0xc9, 0xae } };
static const GUID MY_IID_IHttpBaseProtocolFilter3= { 0xd43f4d4c, 0xbd42, 0x43ae, { 0x87, 0x17, 0xad, 0x2c, 0x8f, 0x4b, 0x29, 0x37 } };
static const GUID MY_IID_IHttpBaseProtocolFilter4= { 0x9fe36ccf, 0x2983, 0x4893, { 0x94, 0x1f, 0xeb, 0x51, 0x8c, 0xa8, 0xce, 0xf9 } };
static const GUID MY_IID_IHttpBaseProtocolFilter5= { 0x416e4993, 0x31e3, 0x4816, { 0xbf, 0x09, 0xe0, 0x18, 0xee, 0x8d, 0xc1, 0xf5 } };
static const GUID MY_IID_IHttpFilter             = { 0xa4cb6dd5, 0x0902, 0x439e, { 0xbf, 0xd7, 0xe1, 0x25, 0x52, 0xb1, 0x65, 0xce } };
static const GUID MY_IID_IHttpCacheControl       = { 0xc77e1cb4, 0x3cea, 0x4eb5, { 0xac, 0x85, 0x04, 0xe1, 0x86, 0xe6, 0x3a, 0xb7 } };
static const GUID MY_IID_IHttpClient             = { 0x7fda1151, 0x3574, 0x4880, { 0xa8, 0xba, 0xe6, 0xb1, 0xe0, 0x06, 0x1f, 0x3d } };
static const GUID MY_IID_IHttpClient2            = { 0xcdd83348, 0xe8b7, 0x4cec, { 0xb1, 0xb0, 0xdc, 0x45, 0x5f, 0xe7, 0x2c, 0x92 } };
static const GUID MY_IID_IHttpClient3            = { 0x1172fd01, 0x9899, 0x4194, { 0x96, 0x3f, 0x8f, 0x9d, 0x72, 0xa7, 0xec, 0x15 } };
static const GUID MY_IID_IHttpClientFactory      = { 0xc30c4eca, 0xe3fa, 0x4f99, { 0xaf, 0xb4, 0x63, 0xcc, 0x65, 0x00, 0x94, 0x62 } };
static const GUID MY_IID_IHttpMethod             = { 0x728d4022, 0x700d, 0x4fe0, { 0xaf, 0xa5, 0x40, 0x29, 0x9c, 0x58, 0xdb, 0xfd } };
static const GUID MY_IID_IHttpMethodFactory      = { 0x3c51d10d, 0x36d7, 0x40f8, { 0xa8, 0x6d, 0xe7, 0x59, 0xca, 0xf2, 0xf8, 0x3f } };
static const GUID MY_IID_IHttpMethodStatics      = { 0x64d171f0, 0xd99a, 0x4153, { 0x8d, 0xc6, 0xd6, 0x8c, 0xc4, 0xcc, 0xe3, 0x17 } };
static const GUID MY_IID_IHttpRequestMessageFactory = { 0x5bac994e, 0x3886, 0x412e, { 0xae, 0xc3, 0x52, 0xec, 0x7f, 0x25, 0x61, 0x6f } };

/* an object is a set of interfaces sharing one refcount; each interface pointer is a {vtbl, obj}
 * pair so methods can find their object */
#define W_MAXIF 8
struct wobj;
struct wif { void *const *vtbl; struct wobj *obj; };
struct wobj {
    struct wif ifs[W_MAXIF]; const GUID *iids[W_MAXIF]; int n;
    LONG ref; BOOL immortal; const WCHAR *clsname;
    int kind;                 /* factories: which class they make */
    WCHAR method[24];         /* HttpMethod */
    struct wobj *cache;       /* filter: its HttpCacheControl */
    UINT32 read_behavior, write_behavior;
};
static LONG g_http_logs;
static void http_log(const char *m) { if (InterlockedIncrement(&g_http_logs) <= 30) OutputDebugStringA(m); }

static HRESULT WINAPI w_QueryInterface(struct wif *s, REFIID iid, void **out);
static ULONG WINAPI w_AddRef(struct wif *s) { return s->obj->immortal ? 2 : (ULONG)InterlockedIncrement(&s->obj->ref); }
static ULONG WINAPI w_Release(struct wif *s)
{
    struct wobj *o = s->obj;
    if (o->immortal) return 1;
    LONG r = InterlockedDecrement(&o->ref);
    if (!r) { if (o->cache) w_Release(&o->cache->ifs[0]); HeapFree(GetProcessHeap(), 0, o); }
    return r;
}
static HRESULT WINAPI w_QueryInterface(struct wif *s, REFIID iid, void **out)
{
    struct wobj *o = s->obj;
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &IID_IUnknown) || IsEqualGUID(iid, &MY_IID_IInspectable)) { *out = &o->ifs[0]; w_AddRef(s); return S_OK; }
    for (int i = 0; i < o->n; i++)
        if (IsEqualGUID(iid, o->iids[i])) { *out = &o->ifs[i]; w_AddRef(s); return S_OK; }
    *out = NULL;
    return E_NOINTERFACE;
}
static HRESULT WINAPI w_GetIids(struct wif *s, ULONG *count, GUID **iids) { (void)s; if (count) *count = 0; if (iids) *iids = NULL; return S_OK; }
static HRESULT hstring_of(const WCHAR *str, HSTRING *out)
{
    resolve_winrt();
    if (!out) return E_POINTER;
    *out = NULL;
    return p_WindowsCreateString ? p_WindowsCreateString(str, lstrlenW(str), out) : E_NOTIMPL;
}
static HRESULT WINAPI w_GetRuntimeClassName(struct wif *s, HSTRING *name) { return hstring_of(s->obj->clsname, name); }
static HRESULT WINAPI w_GetTrustLevel(struct wif *s, int *level) { (void)s; if (level) *level = 0; return S_OK; }
#define W_INSPECTABLE (void *)w_QueryInterface, (void *)w_AddRef, (void *)w_Release, (void *)w_GetIids, (void *)w_GetRuntimeClassName, (void *)w_GetTrustLevel

static HRESULT WINAPI w_ok(void) { return S_OK; }
static HRESULT WINAPI w_bool_true(struct wif *s, BYTE *v) { (void)s; if (!v) return E_POINTER; *v = 1; return S_OK; }
static HRESULT WINAPI w_bool_false(struct wif *s, BYTE *v) { (void)s; if (!v) return E_POINTER; *v = 0; return S_OK; }
static HRESULT WINAPI w_u32_zero(struct wif *s, UINT32 *v) { (void)s; if (!v) return E_POINTER; *v = 0; return S_OK; }
static HRESULT WINAPI w_obj_null(struct wif *s, void **v) { (void)s; if (!v) return E_POINTER; *v = NULL; return S_OK; }
static HRESULT WINAPI w_obj_notimpl(struct wif *s, void **v) { (void)s; if (v) *v = NULL; return E_NOTIMPL; }
/* requests: the async operation out parameter is the last one; one, two or three inputs */
static HRESULT WINAPI w_fail1(struct wif *s, void *a, void **op) { (void)s; (void)a; if (op) *op = NULL; http_log("ms365 ole32 shim: Windows.Web.Http request refused (stub)"); return HTTP_E_CANNOT_CONNECT; }
static HRESULT WINAPI w_fail2(struct wif *s, void *a, void *b, void **op) { (void)b; return w_fail1(s, a, op); }

static struct wobj *wobj_new(const WCHAR *clsname, int n, void *const *const *vtbls, const GUID *const *iids)
{
    struct wobj *o = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*o));
    if (!o) return NULL;
    o->ref = 1; o->clsname = clsname; o->n = n;
    for (int i = 0; i < n; i++) { o->ifs[i].vtbl = vtbls[i]; o->ifs[i].obj = o; o->iids[i] = iids[i]; }
    return o;
}

/* HttpCacheControl */
static HRESULT WINAPI cc_get_read(struct wif *s, UINT32 *v) { if (!v) return E_POINTER; *v = s->obj->read_behavior; return S_OK; }
static HRESULT WINAPI cc_put_read(struct wif *s, UINT32 v) { s->obj->read_behavior = v; return S_OK; }
static HRESULT WINAPI cc_get_write(struct wif *s, UINT32 *v) { if (!v) return E_POINTER; *v = s->obj->write_behavior; return S_OK; }
static HRESULT WINAPI cc_put_write(struct wif *s, UINT32 v) { s->obj->write_behavior = v; return S_OK; }
static void *const cachecontrol_vtbl[] = { W_INSPECTABLE, (void *)cc_get_read, (void *)cc_put_read, (void *)cc_get_write, (void *)cc_put_write };

/* HttpBaseProtocolFilter */
static HRESULT WINAPI f_get_cache(struct wif *s, void **v)
{
    if (!v) return E_POINTER;
    struct wobj *o = s->obj;
    if (!o->cache) {
        void *const *vt[] = { cachecontrol_vtbl }; const GUID *ii[] = { &MY_IID_IHttpCacheControl };
        if (!(o->cache = wobj_new(L"Windows.Web.Http.Filters.HttpCacheControl", 1, vt, ii))) { *v = NULL; return E_OUTOFMEMORY; }
    }
    w_AddRef(&o->cache->ifs[0]);
    *v = &o->cache->ifs[0];
    return S_OK;
}
static HRESULT WINAPI f_get_maxconn(struct wif *s, UINT32 *v) { (void)s; if (!v) return E_POINTER; *v = 6; return S_OK; }
static HRESULT WINAPI f_get_maxversion(struct wif *s, UINT32 *v) { (void)s; if (!v) return E_POINTER; *v = 2; /* HttpVersion.Http11 */ return S_OK; }
static HRESULT WINAPI f_add_event(struct wif *s, void *handler, INT64 *token) { (void)s; (void)handler; if (!token) return E_POINTER; *token = 1; return S_OK; }
static void *const filter_vtbl[] = { W_INSPECTABLE,
    (void *)w_bool_true, (void *)w_ok,        /* AllowAutoRedirect */
    (void *)w_bool_false, (void *)w_ok,       /* AllowUI */
    (void *)w_bool_true, (void *)w_ok,        /* AutomaticDecompression */
    (void *)f_get_cache,                      /* CacheControl */
    (void *)w_obj_notimpl,                    /* CookieManager */
    (void *)w_obj_null, (void *)w_ok,         /* ClientCertificate */
    (void *)w_obj_notimpl,                    /* IgnorableServerCertificateErrors */
    (void *)f_get_maxconn, (void *)w_ok,      /* MaxConnectionsPerServer */
    (void *)w_obj_null, (void *)w_ok,         /* ProxyCredential */
    (void *)w_obj_null, (void *)w_ok,         /* ServerCredential */
    (void *)w_bool_true, (void *)w_ok,        /* UseProxy */
};
static void *const filter2_vtbl[] = { W_INSPECTABLE, (void *)f_get_maxversion, (void *)w_ok };
static void *const filter3_vtbl[] = { W_INSPECTABLE, (void *)w_u32_zero, (void *)w_ok };          /* CookieUsageBehavior */
static void *const filter4_vtbl[] = { W_INSPECTABLE, (void *)f_add_event, (void *)w_ok, (void *)w_ok }; /* ServerCustomValidationRequested, ClearAuthenticationCache */
static void *const filter5_vtbl[] = { W_INSPECTABLE, (void *)w_obj_null };                        /* User */
static void *const httpfilter_vtbl[] = { W_INSPECTABLE, (void *)w_fail1 };                        /* SendRequestAsync */
static void *const closable_vtbl[] = { W_INSPECTABLE, (void *)w_ok };
static struct wobj *filter_new(void)
{
    void *const *vt[] = { filter_vtbl, filter2_vtbl, filter3_vtbl, filter4_vtbl, filter5_vtbl, httpfilter_vtbl, closable_vtbl };
    const GUID *ii[] = { &MY_IID_IHttpBaseProtocolFilter, &MY_IID_IHttpBaseProtocolFilter2, &MY_IID_IHttpBaseProtocolFilter3,
                         &MY_IID_IHttpBaseProtocolFilter4, &MY_IID_IHttpBaseProtocolFilter5, &MY_IID_IHttpFilter, &MY_IID_IClosable };
    http_log("ms365 ole32 shim: created Windows.Web.Http.Filters.HttpBaseProtocolFilter (stub)");
    return wobj_new(L"Windows.Web.Http.Filters.HttpBaseProtocolFilter", 7, vt, ii);
}

/* HttpClient */
static HRESULT WINAPI w_tostring(struct wif *s, HSTRING *v) { return hstring_of(s->obj->clsname, v); }
static void *const client_vtbl[] = { W_INSPECTABLE,
    (void *)w_fail1,                          /* DeleteAsync(uri) */
    (void *)w_fail1, (void *)w_fail2,         /* GetAsync(uri), GetAsync(uri, option) */
    (void *)w_fail1, (void *)w_fail1, (void *)w_fail1, /* GetBufferAsync, GetInputStreamAsync, GetStringAsync */
    (void *)w_fail2, (void *)w_fail2,         /* PostAsync(uri, content), PutAsync(uri, content) */
    (void *)w_fail1, (void *)w_fail2,         /* SendRequestAsync(request), SendRequestAsync(request, option) */
    (void *)w_obj_notimpl,                    /* DefaultRequestHeaders */
};
static void *const client2_vtbl[] = { W_INSPECTABLE,
    (void *)w_fail1, (void *)w_fail1, (void *)w_fail2, (void *)w_fail1, (void *)w_fail1, (void *)w_fail1,
    (void *)w_fail2, (void *)w_fail2, (void *)w_fail1, (void *)w_fail2 };
static void *const client3_vtbl[] = { W_INSPECTABLE, (void *)w_obj_null, (void *)w_ok };          /* DefaultPrivacyAnnotation */
static void *const stringable_vtbl[] = { W_INSPECTABLE, (void *)w_tostring };
static struct wobj *client_new(void)
{
    void *const *vt[] = { client_vtbl, client2_vtbl, client3_vtbl, closable_vtbl, stringable_vtbl };
    const GUID *ii[] = { &MY_IID_IHttpClient, &MY_IID_IHttpClient2, &MY_IID_IHttpClient3, &MY_IID_IClosable, &MY_IID_IStringable };
    http_log("ms365 ole32 shim: created Windows.Web.Http.HttpClient (stub)");
    return wobj_new(L"Windows.Web.Http.HttpClient", 5, vt, ii);
}

/* HttpMethod */
static HRESULT WINAPI m_get_method(struct wif *s, HSTRING *v) { return hstring_of(s->obj->method, v); }
static void *const method_vtbl[] = { W_INSPECTABLE, (void *)m_get_method };
static void *const method_stringable_vtbl[] = { W_INSPECTABLE, (void *)m_get_method };
static HRESULT method_new(const WCHAR *name, void **out)
{
    void *const *vt[] = { method_vtbl, method_stringable_vtbl };
    const GUID *ii[] = { &MY_IID_IHttpMethod, &MY_IID_IStringable };
    if (!out) return E_POINTER;
    struct wobj *o = wobj_new(L"Windows.Web.Http.HttpMethod", 2, vt, ii);
    if (!o) { *out = NULL; return E_OUTOFMEMORY; }
    lstrcpynW(o->method, name, sizeof(o->method) / sizeof(WCHAR));
    *out = &o->ifs[0];
    return S_OK;
}

/* factories (one immortal object per class) */
enum { HTTP_FILTER, HTTP_CLIENT, HTTP_METHOD, HTTP_REQUEST, HTTP_NKINDS };
static HRESULT WINAPI fac_ActivateInstance(struct wif *s, void **out)
{
    struct wobj *o = NULL;
    if (!out) return E_POINTER;
    *out = NULL;
    switch (s->obj->kind) {
    case HTTP_FILTER: o = filter_new(); break;
    case HTTP_CLIENT: o = client_new(); break;
    default: return HTTP_E_CANNOT_CONNECT;    /* HttpRequestMessage: requests fail */
    }
    if (!o) return E_OUTOFMEMORY;
    *out = &o->ifs[0];
    return S_OK;
}
static HRESULT WINAPI fac_client_create(struct wif *s, void *filter, void **out)
{
    (void)s; (void)filter;
    struct wobj *o;
    if (!out) return E_POINTER;
    if (!(o = client_new())) { *out = NULL; return E_OUTOFMEMORY; }
    *out = &o->ifs[0];
    return S_OK;
}
static HRESULT WINAPI fac_method_create(struct wif *s, HSTRING name, void **out)
{
    (void)s;
    UINT32 len = 0; const WCHAR *str;
    resolve_winrt();
    str = p_WindowsGetStringRawBuffer ? p_WindowsGetStringRawBuffer(name, &len) : NULL;
    return method_new(str ? str : L"GET", out);
}
static HRESULT WINAPI fac_delete(struct wif *s, void **o) { (void)s; return method_new(L"DELETE", o); }
static HRESULT WINAPI fac_get(struct wif *s, void **o) { (void)s; return method_new(L"GET", o); }
static HRESULT WINAPI fac_head(struct wif *s, void **o) { (void)s; return method_new(L"HEAD", o); }
static HRESULT WINAPI fac_options(struct wif *s, void **o) { (void)s; return method_new(L"OPTIONS", o); }
static HRESULT WINAPI fac_patch(struct wif *s, void **o) { (void)s; return method_new(L"PATCH", o); }
static HRESULT WINAPI fac_post(struct wif *s, void **o) { (void)s; return method_new(L"POST", o); }
static HRESULT WINAPI fac_put(struct wif *s, void **o) { (void)s; return method_new(L"PUT", o); }
static void *const actfactory_vtbl[] = { W_INSPECTABLE, (void *)fac_ActivateInstance };
static void *const clientfactory_vtbl[] = { W_INSPECTABLE, (void *)fac_client_create };
static void *const methodfactory_vtbl[] = { W_INSPECTABLE, (void *)fac_method_create };
static void *const methodstatics_vtbl[] = { W_INSPECTABLE, (void *)fac_delete, (void *)fac_get, (void *)fac_head, (void *)fac_options, (void *)fac_patch, (void *)fac_post, (void *)fac_put };
static void *const requestfactory_vtbl[] = { W_INSPECTABLE, (void *)w_fail2 };                    /* Create(method, uri) */
static struct wobj *g_http_factories[HTTP_NKINDS];

static HRESULT http_stub_factory(const WCHAR *name, REFIID iid, void **out)
{
    static const WCHAR *const names[HTTP_NKINDS] = {
        L"Windows.Web.Http.Filters.HttpBaseProtocolFilter", L"Windows.Web.Http.HttpClient",
        L"Windows.Web.Http.HttpMethod", L"Windows.Web.Http.HttpRequestMessage" };
    int k;
    for (k = 0; k < HTTP_NKINDS; k++) if (lstrcmpW(name, names[k]) == 0) break;
    if (k == HTTP_NKINDS) return REGDB_E_CLASSNOTREG;
    if (!g_http_factories[k]) {
        struct wobj *o = NULL;
        if (k == HTTP_FILTER) { void *const *vt[] = { actfactory_vtbl }; const GUID *ii[] = { &MY_IID_IActivationFactory }; o = wobj_new(names[k], 1, vt, ii); }
        else if (k == HTTP_CLIENT) { void *const *vt[] = { actfactory_vtbl, clientfactory_vtbl }; const GUID *ii[] = { &MY_IID_IActivationFactory, &MY_IID_IHttpClientFactory }; o = wobj_new(names[k], 2, vt, ii); }
        else if (k == HTTP_METHOD) { void *const *vt[] = { actfactory_vtbl, methodfactory_vtbl, methodstatics_vtbl }; const GUID *ii[] = { &MY_IID_IActivationFactory, &MY_IID_IHttpMethodFactory, &MY_IID_IHttpMethodStatics }; o = wobj_new(names[k], 3, vt, ii); }
        else { void *const *vt[] = { actfactory_vtbl, requestfactory_vtbl }; const GUID *ii[] = { &MY_IID_IActivationFactory, &MY_IID_IHttpRequestMessageFactory }; o = wobj_new(names[k], 2, vt, ii); }
        if (!o) return E_OUTOFMEMORY;
        o->immortal = TRUE; o->kind = k;
        if (InterlockedCompareExchangePointer((void **)&g_http_factories[k], o, NULL) != NULL) HeapFree(GetProcessHeap(), 0, o);
        char msg[160]; wsprintfA(msg, "ms365 ole32 shim: serving %ls (stub)", names[k]); http_log(msg);
    }
    return w_QueryInterface(&g_http_factories[k]->ifs[0], iid, out);
}

/* The launcher registers the stand-in classes under WindowsRuntime\ActivatableClassId with this DLL
 * (as ms365shim.dll) as their DllPath, so combase activates them however the caller reaches
 * RoGetActivationFactory (react-native-win32 resolves it at run time through an API set). */
__declspec(dllexport) HRESULT WINAPI DllGetActivationFactory(HSTRING cls, void **factory)
{
    UINT32 len = 0; const WCHAR *name;
    if (!factory) return E_POINTER;
    *factory = NULL;
    resolve_winrt();
    name = p_WindowsGetStringRawBuffer ? p_WindowsGetStringRawBuffer(cls, &len) : NULL;
    if (!name) return CLASS_E_CLASSNOTAVAILABLE;
    HRESULT hr = http_stub_factory(name, &MY_IID_IActivationFactory, factory);
    return hr == REGDB_E_CLASSNOTREG ? CLASS_E_CLASSNOTAVAILABLE : hr;
}

static HRESULT WINAPI my_RoGetActivationFactory(HSTRING cls, REFIID iid, void **out)
{
    resolve_winrt();
    if (!real_RoGetActivationFactory) return E_NOTIMPL;
    HRESULT hr = real_RoGetActivationFactory(cls, iid, out);
    if (SUCCEEDED(hr) || !iid || !out) return hr;
    UINT32 len = 0; const WCHAR *name = p_WindowsGetStringRawBuffer ? p_WindowsGetStringRawBuffer(cls, &len) : NULL;
    if (name && IsEqualGUID(iid, &MY_IID_ILanguageStatics) && lstrcmpW(name, L"Windows.Globalization.Language") == 0) {
        OutputDebugStringA("ms365 ole32 shim: served ILanguageStatics for Windows.Globalization.Language");
        *out = &lang_statics_obj;
        return S_OK;
    }
    if (name && IsEqualGUID(iid, &MY_IID_IJsonObjectStatics) && lstrcmpW(name, L"Windows.Data.Json.JsonObject") == 0) {
        OutputDebugStringA("ms365 ole32 shim: served IJsonObjectStatics for Windows.Data.Json.JsonObject");
        *out = &json_statics_obj;
        return S_OK;
    }
    if (name && wcsncmp(name, L"Windows.Web.Http.", 17) == 0) {
        HRESULT h2 = http_stub_factory(name, iid, out);
        if (h2 != REGDB_E_CLASSNOTREG) return h2;
    }
    return hr;
}


/* ---- Windows Installer: empty product code = "whichever package owns it" ----------------------
 * Office validates its proofing host through msi.dll with an empty product code
 * (MsiQueryFeatureState("", "OfficeMSProof6"), MsiGetComponentPathEx("", msspell7 component)).
 * Click-to-Run's App-V layer answers those on Windows; Wine's msi rejects an empty product as
 * INSTALLSTATE_INVALIDARG and Word decides the speller is missing. Resolve the empty product against
 * the registered products (launcher: msi-components.py) and answer with the first that knows it. */
typedef INSTALLSTATE (WINAPI *pMsiQueryFeatureStateW)(LPCWSTR, LPCWSTR);
typedef INSTALLSTATE (WINAPI *pMsiGetComponentPathExW)(LPCWSTR, LPCWSTR, LPCWSTR, MSIINSTALLCONTEXT, LPWSTR, LPDWORD);
typedef INSTALLSTATE (WINAPI *pMsiGetComponentPathW)(LPCWSTR, LPCWSTR, LPWSTR, LPDWORD);
typedef UINT (WINAPI *pMsiEnumClientsW)(LPCWSTR, DWORD, LPWSTR);
typedef UINT (WINAPI *pMsiEnumProductsW)(DWORD, LPWSTR);
static pMsiQueryFeatureStateW real_MsiQueryFeatureStateW; static pMsiGetComponentPathExW real_MsiGetComponentPathExW;
static pMsiGetComponentPathW real_MsiGetComponentPathW; static pMsiEnumClientsW real_MsiEnumClientsW; static pMsiEnumProductsW real_MsiEnumProductsW;
static LONG g_msi_logs;
static void msi_log(const char *fn, LPCWSTR arg, LPCWSTR product, int state)
{
    char msg[300];
    if (InterlockedIncrement(&g_msi_logs) > 40) return;
    wsprintfA(msg, "ms365 ole32 shim: %s(\"\", %ls) -> product %ls state %d", fn, arg, product ? product : L"(none)", state);
    OutputDebugStringA(msg);
}
static INSTALLSTATE WINAPI my_MsiQueryFeatureStateW(LPCWSTR product, LPCWSTR feature)
{
    WCHAR prod[64];
    if ((product && product[0]) || !real_MsiEnumProductsW || !feature) return real_MsiQueryFeatureStateW(product, feature);
    for (DWORD i = 0; real_MsiEnumProductsW(i, prod) == ERROR_SUCCESS; i++) {
        INSTALLSTATE st = real_MsiQueryFeatureStateW(prod, feature);
        if (st == INSTALLSTATE_LOCAL || st == INSTALLSTATE_SOURCE || st == INSTALLSTATE_ADVERTISED) { msi_log("MsiQueryFeatureStateW", feature, prod, st); return st; }
    }
    msi_log("MsiQueryFeatureStateW", feature, NULL, INSTALLSTATE_UNKNOWN);
    return INSTALLSTATE_UNKNOWN;
}
static INSTALLSTATE component_path_any(LPCWSTR component, LPCWSTR sid, MSIINSTALLCONTEXT ctx, int ex, LPWSTR buf, LPDWORD len, const char *fn)
{
    WCHAR prod[64]; INSTALLSTATE st = INSTALLSTATE_UNKNOWN;
    for (DWORD i = 0; real_MsiEnumClientsW(component, i, prod) == ERROR_SUCCESS; i++) {
        DWORD n = len ? *len : 0;
        st = ex ? real_MsiGetComponentPathExW(prod, component, sid, ctx, buf, len ? &n : NULL) : real_MsiGetComponentPathW(prod, component, buf, len ? &n : NULL);
        if (st == INSTALLSTATE_LOCAL || st == INSTALLSTATE_SOURCE || st == INSTALLSTATE_MOREDATA) { if (len) *len = n; msi_log(fn, component, prod, st); return st; }
    }
    msi_log(fn, component, NULL, st);
    return st;
}
static INSTALLSTATE WINAPI my_MsiGetComponentPathExW(LPCWSTR product, LPCWSTR component, LPCWSTR sid, MSIINSTALLCONTEXT ctx, LPWSTR buf, LPDWORD len)
{
    if ((product && product[0]) || !real_MsiEnumClientsW || !component) return real_MsiGetComponentPathExW(product, component, sid, ctx, buf, len);
    return component_path_any(component, sid, ctx, 1, buf, len, "MsiGetComponentPathExW");
}
static INSTALLSTATE WINAPI my_MsiGetComponentPathW(LPCWSTR product, LPCWSTR component, LPWSTR buf, LPDWORD len)
{
    if ((product && product[0]) || !real_MsiEnumClientsW || !component) return real_MsiGetComponentPathW(product, component, buf, len);
    return component_path_any(component, NULL, 0, 0, buf, len, "MsiGetComponentPathW");
}


/* ---- Office's border-effect windows -----------------------------------------------------------
 * Office draws dialog and menu shadows with unowned, layered MSO_BORDEREFFECT_WINDOW_CLASS popups
 * (thin strips around the decorated window). Wine's Wayland driver can only turn an unowned popup
 * into an independent xdg toplevel, which tiling compositors (sway) then tile: the layout reshuffles,
 * Word's document swapchain loses its surface, and everything flickers until the strip goes away.
 * Give those strips an owner (the active window, i.e. the one being decorated), which makes them
 * transient windows: positioned relative to the owner and floated by the compositor. Only on the
 * Wayland driver: on X11 the owner change makes Office raise its fatal assertion (0xe0000002) a
 * second later, and X11 has no need for it. MS365_OWN_BORDERS=0/1 overrides.
 * Floating still is not enough: a Wayland client cannot place its toplevels, so sway centres each
 * strip on the output and a maximized window's right and bottom shadows become a cross-hair through
 * the middle of the screen. So the strips are also kept hidden: showing a window always goes through
 * SetWindowPos, which sends WM_WINDOWPOSCHANGING before the visibility changes, and the subclassed
 * strip drops SWP_SHOWWINDOW there. Office loses its window shadows, nothing else. */
typedef HWND (WINAPI *pCreateWindowExW)(DWORD, LPCWSTR, LPCWSTR, DWORD, int, int, int, int, HWND, HMENU, HINSTANCE, LPVOID);
static pCreateWindowExW real_CreateWindowExW; static LONG g_border_logs; static int g_own_borders = -1;
static int own_borders(void)
{
    if (g_own_borders < 0) {
        char v[32] = "";
        if (GetEnvironmentVariableA("MS365_OWN_BORDERS", v, sizeof(v))) g_own_borders = v[0] == '1';
        else g_own_borders = GetEnvironmentVariableA("WINE_GRAPHICS_DRIVER", v, sizeof(v)) && lstrcmpiA(v, "wayland") == 0;
    }
    return g_own_borders;
}
/* The window being decorated: the active window if there is one, otherwise the thread's main
 * visible top-level window (under Wayland nothing is active until the compositor hands focus over,
 * and Office creates the strips before that). */
static BOOL CALLBACK find_main_window(HWND hwnd, LPARAM lp)
{
    LONG style = GetWindowLongW(hwnd, GWL_STYLE), ex = GetWindowLongW(hwnd, GWL_EXSTYLE);
    if ((style & WS_VISIBLE) && !(style & WS_CHILD) && !(ex & WS_EX_TOOLWINDOW) && (style & WS_CAPTION) == WS_CAPTION) {
        *(HWND *)lp = hwnd;
        return FALSE;
    }
    return TRUE;
}
static HWND decorated_window(HWND self)
{
    HWND owner = GetActiveWindow();
    if (!owner || owner == self) owner = GetForegroundWindow();
    if (!owner || owner == self) { owner = NULL; EnumThreadWindows(GetCurrentThreadId(), find_main_window, (LPARAM)&owner); }
    return owner != self ? owner : NULL;
}
static const WCHAR BORDER_PROC_PROP[] = L"ms365.borderproc";
static LRESULT CALLBACK border_wndproc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    WNDPROC orig = (WNDPROC)GetPropW(hwnd, BORDER_PROC_PROP);
    LRESULT r = orig ? CallWindowProcW(orig, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    if (msg == WM_WINDOWPOSCHANGING && lp) ((WINDOWPOS *)lp)->flags &= ~SWP_SHOWWINDOW;
    else if (msg == WM_NCDESTROY) RemovePropW(hwnd, BORDER_PROC_PROP);
    return r;
}
static HWND WINAPI my_CreateWindowExW(DWORD ex, LPCWSTR cls, LPCWSTR name, DWORD style, int x, int y, int w, int h, HWND parent, HMENU menu, HINSTANCE inst, LPVOID param)
{
    HWND hwnd = real_CreateWindowExW(ex, cls, name, style, x, y, w, h, parent, menu, inst, param);
    if (hwnd && !parent && (style & WS_POPUP) && (ex & WS_EX_LAYERED) && own_borders()) {
        WCHAR cn[64];
        if (GetClassNameW(hwnd, cn, 64) && lstrcmpiW(cn, L"MSO_BORDEREFFECT_WINDOW_CLASS") == 0) {
            WNDPROC orig = (WNDPROC)GetWindowLongPtrW(hwnd, GWLP_WNDPROC);
            if (orig && orig != border_wndproc && SetPropW(hwnd, BORDER_PROC_PROP, (HANDLE)orig))
                SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)border_wndproc);
            if (style & WS_VISIBLE) ShowWindow(hwnd, SW_HIDE);
            HWND owner = decorated_window(hwnd);
            if (owner) {
                SetWindowLongPtrW(hwnd, GWLP_HWNDPARENT, (LONG_PTR)owner);
                if (InterlockedIncrement(&g_border_logs) <= 20) {
                    char msg[160]; wsprintfA(msg, "ms365 ole32 shim: border-effect window %p owned by %p, kept hidden", hwnd, owner); OutputDebugStringA(msg);
                }
            }
        }
    }
    return hwnd;
}

/* ---- Direct2D SVG documents -------------------------------------------------------------------
 * Office's licensing / activation dialog draws its icons with ID2D1DeviceContext5::CreateSvgDocument.
 * Wine's d2d1 has that as a stub returning E_NOTIMPL; Office does not check and writes through the
 * missing document (access violation in mso20win32client, Word exits with code 64 before the sign-in
 * page can open). The first D2D1CreateFactory call patches the device-context vtable (shared by
 * every context in the process) so CreateSvgDocument falls back to an empty document: it parses
 * nothing, has one root element that accepts any attribute, and draws nothing (Wine's
 * DrawSvgDocument is a no-op). The icons stay blank, the dialog works. */
typedef struct { float width, height; } my_size_f;
static const GUID MY_IID_IUnknown         = { 0x00000000, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
static const GUID MY_IID_ID2D1Resource    = { 0x2cd90691, 0x12e2, 0x11dc, { 0x9f, 0xed, 0x00, 0x11, 0x43, 0xa0, 0x55, 0xf9 } };
static const GUID MY_IID_ID2D1SvgDocument = { 0x86b88e4d, 0xafa4, 0x4d7b, { 0x88, 0xe4, 0x68, 0xa5, 0x1c, 0x4a, 0x0a, 0xec } };
static const GUID MY_IID_ID2D1SvgElement  = { 0xac7b67a6, 0x183e, 0x49c1, { 0xa8, 0x23, 0x0e, 0xbe, 0x40, 0xb0, 0xdb, 0x29 } };
static const GUID MY_IID_ID2D1DeviceContext5 = { 0x7836d248, 0x68cc, 0x4df6, { 0xb9, 0xe8, 0xde, 0x99, 0x1b, 0xf6, 0x2e, 0xb7 } };
#define D2D_CTX_GETFACTORY        3
#define D2D_CTX_CREATESVGDOCUMENT 115   /* ID2D1DeviceContext6 vtable slot (Wine dlls/d2d1/device.c) */
#define D2D_FACTORY_CREATEDCRT    16    /* ID2D1Factory::CreateDCRenderTarget */
#define COM_CALL(obj, idx, type)  ((type)((*(void ***)(obj))[idx]))
typedef ULONG (WINAPI *pUnkRelease)(void *);
typedef HRESULT (WINAPI *pUnkQI)(void *, REFIID, void **);

struct svg_doc;
struct svg_elem { void *const *vtbl; LONG ref; struct svg_doc *doc; };  /* ref unused for the root */
struct svg_doc  { void *const *vtbl; LONG ref; void *factory; my_size_f viewport; struct svg_elem root; };
static LONG g_svg_logs;
static void svg_log(const char *msg) { if (InterlockedIncrement(&g_svg_logs) <= 20) OutputDebugStringA(msg); }

static ULONG WINAPI svgdoc_AddRef(struct svg_doc *d) { return InterlockedIncrement(&d->ref); }
static ULONG WINAPI svgdoc_Release(struct svg_doc *d)
{
    LONG r = InterlockedDecrement(&d->ref);
    if (!r) { if (d->factory) COM_CALL(d->factory, 2, pUnkRelease)(d->factory); HeapFree(GetProcessHeap(), 0, d); }
    return r;
}
static HRESULT WINAPI svgdoc_QueryInterface(struct svg_doc *d, REFIID iid, void **out)
{
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &MY_IID_IUnknown) || IsEqualGUID(iid, &MY_IID_ID2D1Resource) || IsEqualGUID(iid, &MY_IID_ID2D1SvgDocument)) {
        svgdoc_AddRef(d); *out = d; return S_OK;
    }
    *out = NULL; return E_NOINTERFACE;
}
static void WINAPI svg_GetFactory_doc(struct svg_doc *d, void **factory)
{
    if (!factory) return;
    *factory = d->factory;
    if (d->factory) COM_CALL(d->factory, 1, pUnkRelease)(d->factory); /* AddRef shares Release's signature */
}

/* elements: the root lives inside its document and shares its refcount; CreateChild makes
 * standalone ones that hold a document reference */
static ULONG WINAPI svgel_AddRef(struct svg_elem *e) { return e == &e->doc->root ? svgdoc_AddRef(e->doc) : (ULONG)InterlockedIncrement(&e->ref); }
static ULONG WINAPI svgel_Release(struct svg_elem *e)
{
    if (e == &e->doc->root) return svgdoc_Release(e->doc);
    LONG r = InterlockedDecrement(&e->ref);
    if (!r) { svgdoc_Release(e->doc); HeapFree(GetProcessHeap(), 0, e); }
    return r;
}
static HRESULT WINAPI svgel_QueryInterface(struct svg_elem *e, REFIID iid, void **out)
{
    if (!out) return E_POINTER;
    if (IsEqualGUID(iid, &MY_IID_IUnknown) || IsEqualGUID(iid, &MY_IID_ID2D1Resource) || IsEqualGUID(iid, &MY_IID_ID2D1SvgElement)) {
        svgel_AddRef(e); *out = e; return S_OK;
    }
    *out = NULL; return E_NOINTERFACE;
}
static void WINAPI svgel_GetFactory(struct svg_elem *e, void **factory) { svg_GetFactory_doc(e->doc, factory); }
static void *const svgel_vtbl[34];
static struct svg_elem *svgel_new(struct svg_doc *d)
{
    struct svg_elem *e = HeapAlloc(GetProcessHeap(), 0, sizeof(*e));
    if (!e) return NULL;
    e->vtbl = svgel_vtbl; e->ref = 1; e->doc = d; svgdoc_AddRef(d);
    return e;
}
static void WINAPI svgel_GetDocument(struct svg_elem *e, void **doc) { if (doc) { svgdoc_AddRef(e->doc); *doc = e->doc; } }
static void WINAPI svgel_out_null(struct svg_elem *e, void **out) { (void)e; if (out) *out = NULL; }                    /* GetParent/GetFirstChild/GetLastChild */
static HRESULT WINAPI svgel_sibling(struct svg_elem *e, void *child, void **out) { (void)e; (void)child; if (out) *out = NULL; return E_INVALIDARG; }
static HRESULT WINAPI svgel_ok(void) { return S_OK; }                    /* tree edits, attribute/text setters */
static HRESULT WINAPI svgel_fail(void) { return E_INVALIDARG; }          /* getters that fill caller buffers: "not set" */
static UINT32 WINAPI svgel_zero(void) { return 0; }                      /* lengths, counts, BOOL queries */
static BOOL WINAPI svgel_IsAttributeSpecified(struct svg_elem *e, LPCWSTR name, BOOL *inherited) { (void)e; (void)name; if (inherited) *inherited = FALSE; return FALSE; }
static HRESULT WINAPI svgel_CreateChild(struct svg_elem *e, LPCWSTR tag, void **out)
{
    (void)tag;
    if (!out) return E_POINTER;
    *out = svgel_new(e->doc);
    return *out ? S_OK : E_OUTOFMEMORY;
}
/* ID2D1SvgElement (d2d1svg.h). The overloaded Set/GetAttributeValue slots all get the same
 * behaviour, so MSVC's ordering of overloads does not matter. */
static void *const svgel_vtbl[34] = {
    svgel_QueryInterface, svgel_AddRef, svgel_Release, svgel_GetFactory,
    svgel_GetDocument,
    svgel_fail,                  /* GetTagName */
    svgel_zero,                  /* GetTagNameLength */
    svgel_zero,                  /* IsTextContent */
    svgel_out_null,              /* GetParent */
    svgel_zero,                  /* HasChildren */
    svgel_out_null,              /* GetFirstChild */
    svgel_out_null,              /* GetLastChild */
    svgel_sibling,               /* GetPreviousChild */
    svgel_sibling,               /* GetNextChild */
    svgel_ok, svgel_ok, svgel_ok, svgel_ok,   /* InsertChildBefore, AppendChild, ReplaceChild, RemoveChild */
    svgel_CreateChild,
    svgel_IsAttributeSpecified,
    svgel_zero,                  /* GetSpecifiedAttributeCount */
    svgel_fail, svgel_fail,      /* GetSpecifiedAttributeName, GetSpecifiedAttributeNameLength */
    svgel_ok,                    /* RemoveAttribute */
    svgel_ok,                    /* SetTextValue */
    svgel_fail,                  /* GetTextValue */
    svgel_zero,                  /* GetTextValueLength */
    svgel_ok, svgel_ok, svgel_ok,             /* SetAttributeValue x3 */
    svgel_fail, svgel_fail, svgel_fail,       /* GetAttributeValue x3 */
    svgel_fail,                  /* GetAttributeValueLength */
};

static HRESULT WINAPI svgdoc_SetViewportSize(struct svg_doc *d, my_size_f size) { d->viewport = size; return S_OK; }
/* MSVC returns structs from member functions through a hidden pointer after `this` */
static my_size_f *WINAPI svgdoc_GetViewportSize(struct svg_doc *d, my_size_f *ret) { *ret = d->viewport; return ret; }
static HRESULT WINAPI svgdoc_SetRoot(struct svg_doc *d, void *root) { (void)d; (void)root; return S_OK; }
static void WINAPI svgdoc_GetRoot(struct svg_doc *d, void **root) { if (root) { svgdoc_AddRef(d); *root = &d->root; } }
static HRESULT WINAPI svgdoc_FindElementById(struct svg_doc *d, LPCWSTR id, void **out) { (void)d; (void)id; if (out) *out = NULL; return S_OK; }
static HRESULT WINAPI svgdoc_Serialize(struct svg_doc *d, void *stream, void *subtree) { (void)d; (void)stream; (void)subtree; return E_NOTIMPL; }
static HRESULT WINAPI svgdoc_Deserialize(struct svg_doc *d, void *stream, void **subtree)
{
    (void)stream;
    if (!subtree) return E_POINTER;
    *subtree = svgel_new(d);
    return *subtree ? S_OK : E_OUTOFMEMORY;
}
static HRESULT WINAPI svgdoc_out4(struct svg_doc *d, void *a, void *b, void *c, void **out) { (void)d; (void)a; (void)b; (void)c; if (out) *out = NULL; return E_NOTIMPL; }
static HRESULT WINAPI svgdoc_out3(struct svg_doc *d, void *a, UINT32 b, void **out) { (void)d; (void)a; (void)b; if (out) *out = NULL; return E_NOTIMPL; }
static HRESULT WINAPI svgdoc_CreatePathData(struct svg_doc *d, void *a, UINT32 b, void *c, UINT32 e, void **out) { (void)d; (void)a; (void)b; (void)c; (void)e; if (out) *out = NULL; return E_NOTIMPL; }
/* ID2D1SvgDocument (d2d1svg.h) */
static void *const svgdoc_vtbl[15] = {
    svgdoc_QueryInterface, svgdoc_AddRef, svgdoc_Release, svg_GetFactory_doc,
    svgdoc_SetViewportSize, svgdoc_GetViewportSize, svgdoc_SetRoot, svgdoc_GetRoot,
    svgdoc_FindElementById, svgdoc_Serialize, svgdoc_Deserialize,
    svgdoc_out4,                 /* CreatePaint(type, color, id, paint) */
    svgdoc_out3,                 /* CreateStrokeDashArray(dashes, count, array) */
    svgdoc_out3,                 /* CreatePointCollection(points, count, collection) */
    svgdoc_CreatePathData,
};

typedef HRESULT (WINAPI *pCreateSvgDocument)(void *, void *, my_size_f, void **);
static pCreateSvgDocument real_CreateSvgDocument;
static HRESULT WINAPI my_CreateSvgDocument(void *ctx, void *stream, my_size_f viewport, void **out)
{
    HRESULT hr = real_CreateSvgDocument ? real_CreateSvgDocument(ctx, stream, viewport, out) : E_NOTIMPL;
    if (hr != E_NOTIMPL || !out) return hr;   /* a Wine that implements SVG answers for itself */
    struct svg_doc *d = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*d));
    if (!d) { *out = NULL; return E_OUTOFMEMORY; }
    d->vtbl = svgdoc_vtbl; d->ref = 1; d->viewport = viewport;
    d->root.vtbl = svgel_vtbl; d->root.doc = d;
    typedef void (WINAPI *pGetFactory)(void *, void **);
    COM_CALL(ctx, D2D_CTX_GETFACTORY, pGetFactory)(ctx, &d->factory);   /* keeps the reference */
    svg_log("ms365 ole32 shim: CreateSvgDocument answered with an empty SVG document");
    *out = d;
    return S_OK;
}

/* ---- Direct2D: D2D1_UNIT_MODE_PIXELS -----------------------------------------------------------------
 * In pixel unit mode a device context takes coordinates, sizes and em sizes as pixels, whatever its DPI.
 * Wine's d2d1 (through at least 11.16) stores the mode but always scales by dpi / 96. Office renders its
 * symbol-font icons (formula bar, sheet-tab splitter) in pixel mode on contexts whose DPI it sets to the
 * system DPI, so under Wine they come out dpi/96 times too large and too far from the origin: shifted
 * down and cut off by their boxes. While a context is in pixel mode, keep Wine's DPI at 96 and report the
 * DPI the application set; the real unit mode (GetUnitMode) is the source of truth, the table below only
 * remembers the application's DPI. MS365_NO_PIXEL_UNITS=1 turns this off. */
#define D2D_RT_RESTOREDRAWINGSTATE 44
#define D2D_RT_SETDPI              51
#define D2D_RT_GETDPI              52
#define D2D_CTX_SETUNITMODE        80
#define D2D_CTX_GETUNITMODE        81
#define MY_UNIT_MODE_PIXELS        1
typedef void (WINAPI *pSetDpi)(void *, float, float);
typedef void (WINAPI *pGetDpi)(void *, float *, float *);
typedef void (WINAPI *pSetUnitMode)(void *, int);
typedef int (WINAPI *pGetUnitMode)(void *);
typedef void (WINAPI *pRestoreDrawingState)(void *, void *);
static pSetDpi real_SetDpi; static pGetDpi real_GetDpi;
static pSetUnitMode real_SetUnitMode; static pGetUnitMode real_GetUnitMode;
static pRestoreDrawingState real_RestoreDrawingState;
static struct { void *ctx; float x, y; } g_px_dpi[256];
static SRWLOCK g_px_lock = SRWLOCK_INIT;
static LONG g_px_logs;

static void px_store(void *ctx, float x, float y)
{
    int free_slot = -1;
    AcquireSRWLockExclusive(&g_px_lock);
    for (int i = 0; i < 256; i++) {
        if (g_px_dpi[i].ctx == ctx) { free_slot = i; break; }
        if (!g_px_dpi[i].ctx && free_slot < 0) free_slot = i;
    }
    if (free_slot >= 0) { g_px_dpi[free_slot].ctx = ctx; g_px_dpi[free_slot].x = x; g_px_dpi[free_slot].y = y; }
    ReleaseSRWLockExclusive(&g_px_lock);
}
static BOOL px_lookup(void *ctx, float *x, float *y, BOOL remove)
{
    BOOL found = FALSE;
    AcquireSRWLockExclusive(&g_px_lock);
    for (int i = 0; i < 256; i++) if (g_px_dpi[i].ctx == ctx) {
        *x = g_px_dpi[i].x; *y = g_px_dpi[i].y; found = TRUE;
        if (remove) g_px_dpi[i].ctx = NULL;
        break;
    }
    ReleaseSRWLockExclusive(&g_px_lock);
    return found;
}
static void px_log(const char *what, void *ctx, int mode, float x)
{
    if (InterlockedIncrement(&g_px_logs) <= 20) {
        char msg[160];
        wsprintfA(msg, "ms365 ole32 shim: %s ctx %p unit mode %d app dpi %d", what, ctx, mode, (int)x);
        OutputDebugStringA(msg);
    }
}
/* the context just switched modes (the real unit mode is already the new one) */
static void px_enter(void *ctx)
{
    float x, y;
    real_GetDpi(ctx, &x, &y);
    px_store(ctx, x, y);
    real_SetDpi(ctx, 96.0f, 96.0f);
    px_log("pixel unit mode on", ctx, MY_UNIT_MODE_PIXELS, x);
}
static void px_leave(void *ctx)
{
    float x, y;
    if (px_lookup(ctx, &x, &y, TRUE)) real_SetDpi(ctx, x, y);
}
static void WINAPI my_SetUnitMode(void *ctx, int mode)
{
    int old = real_GetUnitMode(ctx);
    real_SetUnitMode(ctx, mode);
    mode = real_GetUnitMode(ctx);
    if (mode == MY_UNIT_MODE_PIXELS && old != MY_UNIT_MODE_PIXELS) px_enter(ctx);
    else if (mode != MY_UNIT_MODE_PIXELS && old == MY_UNIT_MODE_PIXELS) px_leave(ctx);
}
static void WINAPI my_RestoreDrawingState(void *ctx, void *block)
{
    int old = real_GetUnitMode(ctx), mode;
    real_RestoreDrawingState(ctx, block);
    mode = real_GetUnitMode(ctx);
    if (mode == MY_UNIT_MODE_PIXELS && old != MY_UNIT_MODE_PIXELS) px_enter(ctx);
    else if (mode != MY_UNIT_MODE_PIXELS && old == MY_UNIT_MODE_PIXELS) px_leave(ctx);
}
static void WINAPI my_SetDpi(void *ctx, float x, float y)
{
    if (real_GetUnitMode(ctx) != MY_UNIT_MODE_PIXELS) { real_SetDpi(ctx, x, y); return; }
    if (x == 0.0f && y == 0.0f) x = y = 96.0f;
    if (x <= 0.0f || y <= 0.0f) { real_SetDpi(ctx, x, y); return; }   /* let Wine reject it */
    px_store(ctx, x, y);
    px_log("SetDpi in pixel unit mode", ctx, MY_UNIT_MODE_PIXELS, x);
}
static void WINAPI my_GetDpi(void *ctx, float *x, float *y)
{
    if (real_GetUnitMode(ctx) == MY_UNIT_MODE_PIXELS && x && y && px_lookup(ctx, x, y, FALSE)) return;
    real_GetDpi(ctx, x, y);
}
static void patch_unit_mode(void *ctx)
{
    void **vt = *(void ***)ctx; DWORD old; char v[4] = "";
    if (GetEnvironmentVariableA("MS365_NO_PIXEL_UNITS", v, sizeof(v)) && v[0] == '1') return;
    if (vt[D2D_CTX_SETUNITMODE] == (void *)my_SetUnitMode) return;
    if (!VirtualProtect(vt, (D2D_CTX_GETUNITMODE + 1) * sizeof(void *), PAGE_READWRITE, &old)) return;
    real_SetDpi = (pSetDpi)vt[D2D_RT_SETDPI];
    real_GetDpi = (pGetDpi)vt[D2D_RT_GETDPI];
    real_SetUnitMode = (pSetUnitMode)vt[D2D_CTX_SETUNITMODE];
    real_GetUnitMode = (pGetUnitMode)vt[D2D_CTX_GETUNITMODE];
    real_RestoreDrawingState = (pRestoreDrawingState)vt[D2D_RT_RESTOREDRAWINGSTATE];
    vt[D2D_RT_SETDPI] = (void *)my_SetDpi;
    vt[D2D_RT_GETDPI] = (void *)my_GetDpi;
    vt[D2D_CTX_SETUNITMODE] = (void *)my_SetUnitMode;
    vt[D2D_RT_RESTOREDRAWINGSTATE] = (void *)my_RestoreDrawingState;
    VirtualProtect(vt, (D2D_CTX_GETUNITMODE + 1) * sizeof(void *), old, &old);
    svg_log("ms365 ole32 shim: ID2D1DeviceContext pixel unit mode emulated");
}

/* ---- Direct2D: axis-aligned clips at non-96 DPI --------------------------------------------------------
 * Wine's d2d1 (through at least 11.16) turns a PushAxisAlignedClip rectangle into device pixels as
 * (rect * dpi/96) * world transform, while everything it draws goes through (point * world transform)
 * * dpi/96: the transform's translation is not scaled for clips. At any DPI other than 96 a translated
 * clip lands off by translation * (dpi/96 - 1). OneNote renders every line of page text into a cache
 * cell that way (clip, clear, clip, white clear, glyphs, at DPI 180 with a translation), so the lines
 * were clipped away and came out as black boxes. While Wine computes the clip, hand it a transform whose
 * translation is already scaled. Command-list targets record the calls instead and are left alone. */
static const GUID MY_IID_ID2D1CommandList = { 0xb4f34a19, 0x2383, 0x4d76, { 0x94, 0xf6, 0xec, 0x34, 0x36, 0x57, 0xc3, 0xdc } };
#define D2D_RT_SETTRANSFORM        30
#define D2D_RT_GETTRANSFORM        31
#define D2D_RT_PUSHAXISALIGNEDCLIP 45
#define D2D_CTX_GETTARGET          75
typedef struct { float _11, _12, _21, _22, _31, _32; } my_mat3x2;
typedef void (WINAPI *pPushAxisAlignedClip)(void *, const void *, int);
typedef void (WINAPI *pRtSetTransform)(void *, const my_mat3x2 *);
typedef void (WINAPI *pRtGetTransform)(void *, my_mat3x2 *);
static pPushAxisAlignedClip real_PushAxisAlignedClip;
static pRtSetTransform real_RtSetTransform; static pRtGetTransform real_RtGetTransform;
static pGetDpi clip_GetDpi;     /* Wine's own GetDpi (the pixel-unit-mode wrapper reports the app's DPI) */
static LONG g_clip_logs;

static BOOL target_is_command_list(void *ctx)
{
    typedef void (WINAPI *pGetTarget)(void *, void **);
    void *target = NULL, *list = NULL; BOOL ret = FALSE;
    COM_CALL(ctx, D2D_CTX_GETTARGET, pGetTarget)(ctx, &target);
    if (!target) return FALSE;
    if (SUCCEEDED(COM_CALL(target, 0, pUnkQI)(target, &MY_IID_ID2D1CommandList, &list)) && list) {
        ret = TRUE;
        COM_CALL(list, 2, pUnkRelease)(list);
    }
    COM_CALL(target, 2, pUnkRelease)(target);
    return ret;
}
static void WINAPI my_PushAxisAlignedClip(void *ctx, const void *rect, int antialias)
{
    float dx = 96.0f, dy = 96.0f; my_mat3x2 m, t;
    clip_GetDpi(ctx, &dx, &dy);
    if ((dx == 96.0f && dy == 96.0f) || !rect) { real_PushAxisAlignedClip(ctx, rect, antialias); return; }
    real_RtGetTransform(ctx, &m);
    if ((m._31 == 0.0f && m._32 == 0.0f) || target_is_command_list(ctx)) { real_PushAxisAlignedClip(ctx, rect, antialias); return; }
    t = m;
    t._31 *= dx / 96.0f;
    t._32 *= dy / 96.0f;
    real_RtSetTransform(ctx, &t);
    real_PushAxisAlignedClip(ctx, rect, antialias);
    real_RtSetTransform(ctx, &m);
    if (InterlockedIncrement(&g_clip_logs) <= 3) OutputDebugStringA("ms365 ole32 shim: translated clip at non-96 DPI corrected");
}
static void patch_axis_aligned_clip(void *ctx)
{
    void **vt = *(void ***)ctx; DWORD old;
    if (vt[D2D_RT_PUSHAXISALIGNEDCLIP] == (void *)my_PushAxisAlignedClip) return;
    if (!VirtualProtect(vt, (D2D_CTX_GETTARGET + 1) * sizeof(void *), PAGE_READWRITE, &old)) return;
    real_PushAxisAlignedClip = (pPushAxisAlignedClip)vt[D2D_RT_PUSHAXISALIGNEDCLIP];
    real_RtSetTransform = (pRtSetTransform)vt[D2D_RT_SETTRANSFORM];
    real_RtGetTransform = (pRtGetTransform)vt[D2D_RT_GETTRANSFORM];
    clip_GetDpi = (pGetDpi)vt[D2D_RT_GETDPI];
    vt[D2D_RT_PUSHAXISALIGNEDCLIP] = (void *)my_PushAxisAlignedClip;
    VirtualProtect(vt, (D2D_CTX_GETTARGET + 1) * sizeof(void *), old, &old);
    svg_log("ms365 ole32 shim: ID2D1RenderTarget::PushAxisAlignedClip DPI fix installed");
}

typedef HRESULT (WINAPI *pD2D1CreateFactory)(int, REFIID, const void *, void **);
static pD2D1CreateFactory real_D2D1CreateFactory;
static LONG g_svg_vtbl_state;
static void patch_svg_vtable(void *factory)
{
    if (InterlockedCompareExchange(&g_svg_vtbl_state, 1, 0) != 0) return;
    /* any device context will do: a DC render target owns one (D2D1_RENDER_TARGET_PROPERTIES:
     * default type, B8G8R8A8_UNORM premultiplied, default dpi/usage/feature level) */
    struct { int type, format, alpha; float dpix, dpiy; int usage, minlevel; } props = { 0, 87, 1, 0.0f, 0.0f, 0, 0 };
    typedef HRESULT (WINAPI *pCreateDCRenderTarget)(void *, const void *, void **);
    void *rt = NULL, *ctx = NULL;
    HRESULT hr = COM_CALL(factory, D2D_FACTORY_CREATEDCRT, pCreateDCRenderTarget)(factory, &props, &rt);
    if (FAILED(hr) || !rt) { svg_log("ms365 ole32 shim: SVG patch: CreateDCRenderTarget failed"); return; }
    if (SUCCEEDED(COM_CALL(rt, 0, pUnkQI)(rt, &MY_IID_ID2D1DeviceContext5, &ctx)) && ctx) {
        void **slot = &(*(void ***)ctx)[D2D_CTX_CREATESVGDOCUMENT];
        DWORD old;
        if (*slot != (void *)my_CreateSvgDocument && VirtualProtect(slot, sizeof(void *), PAGE_READWRITE, &old)) {
            real_CreateSvgDocument = (pCreateSvgDocument)*slot;
            *slot = (void *)my_CreateSvgDocument;
            VirtualProtect(slot, sizeof(void *), old, &old);
            svg_log("ms365 ole32 shim: SVG patch: ID2D1DeviceContext5::CreateSvgDocument wrapped");
        }
        patch_axis_aligned_clip(ctx);   /* first: it keeps Wine's own GetDpi */
        patch_unit_mode(ctx);
        COM_CALL(ctx, 2, pUnkRelease)(ctx);
    } else svg_log("ms365 ole32 shim: SVG patch: no ID2D1DeviceContext5");
    COM_CALL(rt, 2, pUnkRelease)(rt);
}
static HRESULT WINAPI my_D2D1CreateFactory(int type, REFIID iid, const void *opts, void **out)
{
    HRESULT hr = real_D2D1CreateFactory ? real_D2D1CreateFactory(type, iid, opts, out) : E_NOTIMPL;
    if (SUCCEEDED(hr) && out && *out) patch_svg_vtable(*out);
    return hr;
}

/* ---- oleaut32: vtable size of 32-bit type libraries in a 64-bit process --------------------
 * Excel's, Word's and Office's type libraries are SYS_WIN32 (they are shared with 32-bit Office)
 * and store per-interface vtable sizes in 4-byte slots: _Worksheet, 158 slots, is stored as 632.
 * Wine's typelib reader scales every method's oVft to native pointers (Range is at 800) but hands
 * TYPEATTR.cbSizeVft out unscaled, so an interface looks shorter than its own last method. VBA
 * sizes the dispatch tables of document modules (ThisWorkbook, Sheet1) from cbSizeVft, and any
 * call through them past slot 79 jumps into garbage: "Automation error" or a crash. This wraps
 * ITypeInfo::GetTypeAttr on Wine's shared type-info vtable, patched when the first type library
 * is loaded, and scales cbSizeVft of interfaces from 32-bit libraries. (Wine reports libraries
 * created in-process with ICreateTypeLib2 in the library's own pointer size as Windows does;
 * MSForms creates only dispinterfaces that way, which this leaves alone.) */
typedef HRESULT (WINAPI *pLoadTypeLibEx)(LPCOLESTR, REGKIND, ITypeLib **);
typedef HRESULT (WINAPI *pLoadTypeLib)(LPCOLESTR, ITypeLib **);
typedef HRESULT (WINAPI *pLoadRegTypeLib)(REFGUID, WORD, WORD, LCID, ITypeLib **);
typedef HRESULT (WINAPI *pGetTypeAttr)(ITypeInfo *, TYPEATTR **);
static pLoadTypeLibEx real_LoadTypeLibEx;
static pLoadTypeLib real_LoadTypeLib;
static pLoadRegTypeLib real_LoadRegTypeLib;
static pGetTypeAttr real_GetTypeAttr;
static HRESULT WINAPI my_GetTypeAttr(ITypeInfo *ti, TYPEATTR **out)
{
    HRESULT hr = real_GetTypeAttr(ti, out);
    ITypeLib *tl; UINT idx; TLIBATTR *la;
    if (FAILED(hr) || !out || !*out || (*out)->typekind != TKIND_INTERFACE || !(*out)->cbSizeVft) return hr;
    if (FAILED(ITypeInfo_GetContainingTypeLib(ti, &tl, &idx))) return hr;
    if (SUCCEEDED(ITypeLib_GetLibAttr(tl, &la))) {
        if (la->syskind != SYS_WIN64) (*out)->cbSizeVft *= sizeof(void *) / 4;
        ITypeLib_ReleaseTLibAttr(tl, la);
    }
    ITypeLib_Release(tl);
    return hr;
}
static void patch_typeinfo_vtable(ITypeLib *tl)
{
    static LONG done; ITypeInfo *ti; void **vt; DWORD old;
    if (done || !tl || !ITypeLib_GetTypeInfoCount(tl) || FAILED(ITypeLib_GetTypeInfo(tl, 0, &ti))) return;
    vt = *(void ***)ti;   /* slot 3: IUnknown's three, then GetTypeAttr */
    if (vt[3] != (void *)my_GetTypeAttr && VirtualProtect(&vt[3], sizeof(void *), PAGE_READWRITE, &old)) {
        real_GetTypeAttr = (pGetTypeAttr)vt[3];
        vt[3] = (void *)my_GetTypeAttr;
        VirtualProtect(&vt[3], sizeof(void *), old, &old);
        done = 1;
        OutputDebugStringA("ms365 ole32 shim: ITypeInfo::GetTypeAttr wrapped (cbSizeVft of 32-bit type libraries)");
    }
    ITypeInfo_Release(ti);
}
static HRESULT WINAPI my_LoadTypeLibEx(LPCOLESTR file, REGKIND kind, ITypeLib **out)
{
    HRESULT hr = real_LoadTypeLibEx ? real_LoadTypeLibEx(file, kind, out) : E_NOTIMPL;
    if (SUCCEEDED(hr) && out) patch_typeinfo_vtable(*out);
    return hr;
}
static HRESULT WINAPI my_LoadTypeLib(LPCOLESTR file, ITypeLib **out)
{
    HRESULT hr = real_LoadTypeLib ? real_LoadTypeLib(file, out) : E_NOTIMPL;
    if (SUCCEEDED(hr) && out) patch_typeinfo_vtable(*out);
    return hr;
}
static HRESULT WINAPI my_LoadRegTypeLib(REFGUID guid, WORD maj, WORD min, LCID lcid, ITypeLib **out)
{
    HRESULT hr = real_LoadRegTypeLib ? real_LoadRegTypeLib(guid, maj, min, lcid, out) : E_NOTIMPL;
    if (SUCCEEDED(hr) && out) patch_typeinfo_vtable(*out);
    return hr;
}

static const struct patch PATCHES[] = {
    { "kernel32.dll", "GetProcAddress",              (void *)my_GetProcAddress },
    { "winhttp.dll",  "WinHttpSetOption",            (void *)my_WinHttpSetOption },
    { "winhttp.dll",  "WinHttpQueryOption",          (void *)my_WinHttpQueryOption },
    { "wininet.dll",  "InternetSetOptionW",          (void *)my_InternetSetOptionW },
    { "wininet.dll",  "InternetSetOptionA",          (void *)my_InternetSetOptionA },
    { "wininet.dll",  "InternetQueryOptionW",        (void *)my_InternetQueryOptionW },
    { "wininet.dll",  "InternetQueryOptionA",        (void *)my_InternetQueryOptionA },
    { "combase.dll",  "RoGetActivationFactory",      (void *)my_RoGetActivationFactory },
    { "msi.dll",      "MsiQueryFeatureStateW",       (void *)my_MsiQueryFeatureStateW },
    { "msi.dll",      "MsiGetComponentPathExW",      (void *)my_MsiGetComponentPathExW },
    { "msi.dll",      "MsiGetComponentPathW",        (void *)my_MsiGetComponentPathW },
    { "user32.dll",   "CreateWindowExW",             (void *)my_CreateWindowExW },
    { "api-ms-win-core-winrt-l1-1-0.dll", "RoGetActivationFactory", (void *)my_RoGetActivationFactory },
    { "kernel32.dll", "SetFileShortNameW",           (void *)my_SetFileShortNameW },
    { "kernel32.dll", "SetFileShortNameA",           (void *)my_SetFileShortNameA },
    { "kernel32.dll", "FindPackagesByPackageFamily", (void *)my_FindPackagesByPackageFamily },
    { "kernel32.dll", "SetThreadpoolTimerEx",        (void *)my_SetThreadpoolTimerEx },
    { "kernel32.dll", "GetDllDirectoryA",            (void *)my_GetDllDirectoryA },
    { "kernel32.dll", "GetDllDirectoryW",            (void *)my_GetDllDirectoryW },
    { "d2d1.dll",     "D2D1CreateFactory",           (void *)my_D2D1CreateFactory },
    { "crypt32.dll",  "CertGetCertificateChain",     (void *)my_CertGetCertificateChain },
    { "virtdisk.dll", "GetStorageDependencyInformation", (void *)my_GetStorageDependencyInformation },
    { "oleaut32.dll", "LoadTypeLibEx",               (void *)my_LoadTypeLibEx },
    { "oleaut32.dll", "LoadTypeLib",                 (void *)my_LoadTypeLib },
    { "oleaut32.dll", "LoadRegTypeLib",              (void *)my_LoadRegTypeLib },
};
#define NPATCHES (sizeof(PATCHES) / sizeof(PATCHES[0]))

/* PATCHES entries that wrap an existing export (not GetProcAddress itself), keyed by module.
 * No API calls in here: Office's App-V layer (AppVIsvSubsystems64) detours kernel32 entry points
 * such as GetModuleHandleA and fails fast when they are re-entered from inside a GetProcAddress
 * that it is making itself. Module handles are recorded by the loader notification instead. */
static struct wrapmod { const char *dll; HMODULE h; } WRAPMODS[] = {
    { "winhttp.dll", NULL },
    { "wininet.dll", NULL },
    { "combase.dll", NULL },
    { "api-ms-win-core-winrt-l1-1-0.dll", NULL },
    { "msi.dll", NULL },
    { "user32.dll", NULL },
    { "d2d1.dll", NULL },
    { "crypt32.dll", NULL },
    { "virtdisk.dll", NULL },
    { "oleaut32.dll", NULL },
};
#define NWRAPMODS (sizeof(WRAPMODS) / sizeof(WRAPMODS[0]))
/* Office imports msi.dll by ordinal (Windows' msi.dll exports MsiQueryFeatureStateW as #111);
 * Wine's msi.spec keeps the same ordinals. Map the ones we wrap back to names. */
static const char *ordinal_name(const char *dll, WORD ordinal)
{
    if (lstrcmpiA(dll, "msi.dll") == 0) {
        if (ordinal == 111) return "MsiQueryFeatureStateW";
        if (ordinal == 173) return "MsiGetComponentPathW";
        if (ordinal == 294) return "MsiGetComponentPathExW";
    }
    /* react-native-win32 (the licensing dialog) links d2d1 #1, D2D1CreateFactory in both d2d1s */
    if (lstrcmpiA(dll, "d2d1.dll") == 0 && ordinal == 1) return "D2D1CreateFactory";
    /* Office links all of oleaut32 by ordinal; Wine's oleaut32.spec keeps Windows' numbering */
    if (lstrcmpiA(dll, "oleaut32.dll") == 0) {
        if (ordinal == 161) return "LoadTypeLib";
        if (ordinal == 162) return "LoadRegTypeLib";
        if (ordinal == 183) return "LoadTypeLibEx";
    }
    return NULL;
}
static const struct patch *wrapped_lookup(HMODULE h, LPCSTR name)
{
    for (size_t m = 0; m < NWRAPMODS; m++) {
        if (!WRAPMODS[m].h || WRAPMODS[m].h != h) continue;
        for (size_t i = 0; i < NPATCHES; i++)
            if (strcmp(PATCHES[i].fn, name) == 0 && lstrcmpiA(PATCHES[i].dll, WRAPMODS[m].dll) == 0) return &PATCHES[i];
        return NULL;
    }
    return NULL;
}

static const struct patch *wrapped_lookup_ordinal(HMODULE h, WORD ordinal)
{
    for (size_t m = 0; m < NWRAPMODS; m++) {
        if (!WRAPMODS[m].h || WRAPMODS[m].h != h) continue;
        const char *nm = ordinal_name(WRAPMODS[m].dll, ordinal);
        return nm ? wrapped_lookup(h, nm) : NULL;
    }
    return NULL;
}

/* ---- machinery --------------------------------------------------------------------------- */

static void dbg(const char *what, const char *mod, const char *fn)
{
    char buf[320];
    wsprintfA(buf, "ms365 ole32 shim: %s %s!%s", what, mod, fn);
    OutputDebugStringA(buf);
}

static int ieq(const char *a, const char *b) { return lstrcmpiA(a, b) == 0; }

/* Rewrite import slots of one module. */
static void note_module(HMODULE h, const char *modname)
{
    for (size_t m = 0; m < NWRAPMODS; m++) {
        if (lstrcmpiA(modname, WRAPMODS[m].dll) != 0) continue;
        WRAPMODS[m].h = h;
        if (m == 0) {
            real_WinHttpSetOption = (pWinHttpSetOption)real_GetProcAddress(h, "WinHttpSetOption");
            real_WinHttpQueryOption = (pWinHttpQueryOption)real_GetProcAddress(h, "WinHttpQueryOption");
        } else if (m == 1) {
            real_InternetSetOptionW = (pInternetSetOption)real_GetProcAddress(h, "InternetSetOptionW");
            real_InternetSetOptionA = (pInternetSetOption)real_GetProcAddress(h, "InternetSetOptionA");
            real_InternetQueryOptionW = (pInternetQueryOption)real_GetProcAddress(h, "InternetQueryOptionW");
            real_InternetQueryOptionA = (pInternetQueryOption)real_GetProcAddress(h, "InternetQueryOptionA");
        } else if (lstrcmpiA(WRAPMODS[m].dll, "user32.dll") == 0) {
            real_CreateWindowExW = (pCreateWindowExW)real_GetProcAddress(h, "CreateWindowExW");
        } else if (lstrcmpiA(WRAPMODS[m].dll, "d2d1.dll") == 0) {
            real_D2D1CreateFactory = (pD2D1CreateFactory)real_GetProcAddress(h, "D2D1CreateFactory");
        } else if (lstrcmpiA(WRAPMODS[m].dll, "crypt32.dll") == 0) {
            real_CertGetCertificateChain = (pCertGetCertificateChain)real_GetProcAddress(h, "CertGetCertificateChain");
        } else if (lstrcmpiA(WRAPMODS[m].dll, "oleaut32.dll") == 0) {
            real_LoadTypeLibEx = (pLoadTypeLibEx)real_GetProcAddress(h, "LoadTypeLibEx");
            real_LoadTypeLib = (pLoadTypeLib)real_GetProcAddress(h, "LoadTypeLib");
            real_LoadRegTypeLib = (pLoadRegTypeLib)real_GetProcAddress(h, "LoadRegTypeLib");
        } else if (lstrcmpiA(WRAPMODS[m].dll, "msi.dll") == 0) {
            real_MsiQueryFeatureStateW = (pMsiQueryFeatureStateW)real_GetProcAddress(h, "MsiQueryFeatureStateW");
            real_MsiGetComponentPathExW = (pMsiGetComponentPathExW)real_GetProcAddress(h, "MsiGetComponentPathExW");
            real_MsiGetComponentPathW = (pMsiGetComponentPathW)real_GetProcAddress(h, "MsiGetComponentPathW");
            real_MsiEnumClientsW = (pMsiEnumClientsW)real_GetProcAddress(h, "MsiEnumClientsW");
            real_MsiEnumProductsW = (pMsiEnumProductsW)real_GetProcAddress(h, "MsiEnumProductsW");
        }
    }
}

static void patch_module(HMODULE h, const char *modname)
{
    BYTE *base = (BYTE *)h;
    IMAGE_DOS_HEADER *dos = (IMAGE_DOS_HEADER *)base;
    if (dos->e_magic != IMAGE_DOS_SIGNATURE) return;
    IMAGE_NT_HEADERS *nt = (IMAGE_NT_HEADERS *)(base + dos->e_lfanew);
    if (nt->Signature != IMAGE_NT_SIGNATURE) return;
    IMAGE_DATA_DIRECTORY dd = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (!dd.VirtualAddress || !dd.Size) return;
    IMAGE_IMPORT_DESCRIPTOR *desc = (IMAGE_IMPORT_DESCRIPTOR *)(base + dd.VirtualAddress);
    for (; desc->Name; desc++) {
        const char *dll = (const char *)(base + desc->Name);
        size_t want = 0;
        for (size_t i = 0; i < NPATCHES; i++) if (ieq(dll, PATCHES[i].dll)) want++;
        if (!want) continue;
        IMAGE_THUNK_DATA *ilt = (IMAGE_THUNK_DATA *)(base + (desc->OriginalFirstThunk ? desc->OriginalFirstThunk : desc->FirstThunk));
        IMAGE_THUNK_DATA *iat = (IMAGE_THUNK_DATA *)(base + desc->FirstThunk);
        for (; ilt->u1.AddressOfData; ilt++, iat++) {
            const char *fn;
            if (IMAGE_SNAP_BY_ORDINAL(ilt->u1.Ordinal)) {
                fn = ordinal_name(dll, (WORD)IMAGE_ORDINAL(ilt->u1.Ordinal));
                if (!fn) continue;
            } else {
                IMAGE_IMPORT_BY_NAME *ibn = (IMAGE_IMPORT_BY_NAME *)(base + ilt->u1.AddressOfData);
                fn = (const char *)ibn->Name;
            }
            for (size_t i = 0; i < NPATCHES; i++) {
                if (!ieq(dll, PATCHES[i].dll) || strcmp(fn, PATCHES[i].fn) != 0) continue;
                if ((void *)iat->u1.Function == PATCHES[i].repl) break;
                DWORD old;
                if (!VirtualProtect(&iat->u1.Function, sizeof(void *), PAGE_READWRITE, &old)) { dbg("VirtualProtect failed", modname, fn); break; }
                iat->u1.Function = (ULONGLONG)(ULONG_PTR)PATCHES[i].repl;
                VirtualProtect(&iat->u1.Function, sizeof(void *), old, &old);
                dbg("patched import", modname, fn);
                break;
            }
        }
    }
}

/* Address of a named export, walking the table by hand (GetProcAddress hides Wine stubs). NULL for
 * forwarders and missing names. */
static BYTE *find_export(HMODULE h, const char *name)
{
    BYTE *base = (BYTE *)h;
    IMAGE_NT_HEADERS *nt = (IMAGE_NT_HEADERS *)(base + ((IMAGE_DOS_HEADER *)base)->e_lfanew);
    IMAGE_DATA_DIRECTORY dd = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
    if (!dd.VirtualAddress) return NULL;
    IMAGE_EXPORT_DIRECTORY *ed = (IMAGE_EXPORT_DIRECTORY *)(base + dd.VirtualAddress);
    DWORD *names = (DWORD *)(base + ed->AddressOfNames);
    WORD *ords = (WORD *)(base + ed->AddressOfNameOrdinals);
    DWORD *funcs = (DWORD *)(base + ed->AddressOfFunctions);
    for (DWORD i = 0; i < ed->NumberOfNames; i++) {
        if (strcmp((const char *)(base + names[i]), name) != 0) continue;
        DWORD rva = funcs[ords[i]];
        if (rva >= dd.VirtualAddress && rva < dd.VirtualAddress + dd.Size) return NULL;
        return base + rva;
    }
    return NULL;
}

/* Overwrite exported stubs of one module (if it is in STUBPATCHES). */
static void patch_stubs(HMODULE h, const char *modname)
{
    for (size_t i = 0; i < NSTUBPATCHES; i++) {
        if (!ieq(modname, STUBPATCHES[i].dll)) continue;
        BYTE *fn = find_export(h, STUBPATCHES[i].fn);
        if (!fn) { dbg("stub export not found", modname, STUBPATCHES[i].fn); continue; }
        if (fn[0] == 0x48 && fn[1] == 0xB8 && fn[10] == 0xFF && fn[11] == 0xE0) continue; /* done */
        DWORD old;
        if (!VirtualProtect(fn, 16, PAGE_EXECUTE_READWRITE, &old)) { dbg("VirtualProtect failed", modname, STUBPATCHES[i].fn); continue; }
        fn[0] = 0x48; fn[1] = 0xB8;                         /* mov rax, imm64 */
        memcpy(fn + 2, &STUBPATCHES[i].repl, sizeof(void *));
        fn[10] = 0xFF; fn[11] = 0xE0;                       /* jmp rax */
        VirtualProtect(fn, 16, old, &old);
        FlushInstructionCache(GetCurrentProcess(), fn, 16);
        dbg("patched stub", modname, STUBPATCHES[i].fn);
    }
}

/* ---- exports to hide ------------------------------------------------------------------------
 * Some APIs exist in this Wine but misbehave; Office probes them with GetProcAddress and has a
 * fallback for Windows builds that lack them. Renaming the export in the in-memory name table
 * (last character bumped, so the sorted order GetProcAddress relies on is preserved) makes the
 * probe fail cleanly.
 *  - QueueUserAPC2: mso30win32client queues special user APCs (flag 1); Wine dispatches them with
 *    the wrong calling convention and Office's APC routine raises fatal 0xE0000002. */
static const struct { const char *dll; const char *fn; } HIDE[] = {
    { "kernel32.dll",   "QueueUserAPC2" },
    { "kernelbase.dll", "QueueUserAPC2" },
};

static void hide_export(const char *dllname, const char *name)
{
    HMODULE h = GetModuleHandleA(dllname);
    if (!h) return;
    BYTE *base = (BYTE *)h;
    IMAGE_NT_HEADERS *nt = (IMAGE_NT_HEADERS *)(base + ((IMAGE_DOS_HEADER *)base)->e_lfanew);
    IMAGE_DATA_DIRECTORY dd = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
    if (!dd.VirtualAddress) return;
    IMAGE_EXPORT_DIRECTORY *ed = (IMAGE_EXPORT_DIRECTORY *)(base + dd.VirtualAddress);
    DWORD *names = (DWORD *)(base + ed->AddressOfNames);
    for (DWORD i = 0; i < ed->NumberOfNames; i++) {
        char *n = (char *)(base + names[i]);
        if (strcmp(n, name) != 0) continue;
        size_t len = strlen(n);
        char last = n[len - 1];
        const char *next = (i + 1 < ed->NumberOfNames) ? (const char *)(base + names[i + 1]) : NULL;
        char tmp[128];
        if (len >= sizeof(tmp)) return;
        memcpy(tmp, n, len + 1);
        tmp[len - 1] = last + 1;
        if (next && strcmp(tmp, next) >= 0) { dbg("cannot hide (order)", dllname, name); return; }
        DWORD old;
        if (!VirtualProtect(n, len + 1, PAGE_READWRITE, &old)) { dbg("cannot hide (protect)", dllname, name); return; }
        n[len - 1] = last + 1;
        VirtualProtect(n, len + 1, old, &old);
        dbg("hid export", dllname, name);
        return;
    }
}

/* Loader notification (ntdll.LdrRegisterDllNotification; Wine implements it). */
typedef struct {
    ULONG Flags;
    const UNICODE_STRING *FullDllName;
    const UNICODE_STRING *BaseDllName;
    void *DllBase;
    ULONG SizeOfImage;
} LDR_NOTIFY_DATA;
typedef void (CALLBACK *LDR_NOTIFY_FN)(ULONG reason, const LDR_NOTIFY_DATA *data, void *ctx);
typedef LONG (NTAPI *pLdrRegisterDllNotification)(ULONG flags, LDR_NOTIFY_FN fn, void *ctx, void **cookie);

static void CALLBACK on_dll_notify(ULONG reason, const LDR_NOTIFY_DATA *data, void *ctx)
{
    (void)ctx;
    if (!data || !data->DllBase) return;
    if (reason == 2 /* LDR_DLL_NOTIFICATION_REASON_UNLOADED */) {
        for (size_t m = 0; m < NWRAPMODS; m++) if ((HMODULE)data->DllBase == WRAPMODS[m].h) WRAPMODS[m].h = NULL;
        return;
    }
    if (reason != 1 /* LDR_DLL_NOTIFICATION_REASON_LOADED */) return;
    char name[128] = "?";
    if (data->BaseDllName && data->BaseDllName->Buffer)
        WideCharToMultiByte(CP_ACP, 0, data->BaseDllName->Buffer, data->BaseDllName->Length / 2, name, sizeof(name) - 1, NULL, NULL);
    if (g_trace_mod_name[0] && lstrcmpiA(name, g_trace_mod_name) == 0) g_trace_mod = (HMODULE)data->DllBase;
    note_module((HMODULE)data->DllBase, name);
    patch_module((HMODULE)data->DllBase, name);
    patch_stubs((HMODULE)data->DllBase, name);
}

typedef BOOL (WINAPI *pEnumProcessModules)(HANDLE, HMODULE *, DWORD, DWORD *);
typedef DWORD (WINAPI *pGetModuleBaseNameA)(HANDLE, HMODULE, LPSTR, DWORD);

static void patch_loaded_modules(void)
{
    HMODULE k32 = GetModuleHandleA("kernel32.dll");
    pEnumProcessModules enumMods = (pEnumProcessModules)GetProcAddress(k32, "K32EnumProcessModules");
    pGetModuleBaseNameA baseName = (pGetModuleBaseNameA)GetProcAddress(k32, "K32GetModuleBaseNameA");
    if (!enumMods || !baseName) return;
    HMODULE mods[1024];
    DWORD needed = 0;
    if (!enumMods(GetCurrentProcess(), mods, sizeof(mods), &needed)) return;
    DWORD n = needed / sizeof(HMODULE);
    if (n > 1024) n = 1024;
    for (DWORD i = 0; i < n; i++) {
        char name[128] = "?";
        baseName(GetCurrentProcess(), mods[i], name, sizeof(name));
        note_module(mods[i], name);
        patch_module(mods[i], name);
        patch_stubs(mods[i], name);
    }
}

static void *notify_cookie;

/* The same DLL is loaded twice: as ole32.dll (the forwarder, whenever Office first calls ole32)
 * and as ms365shim.dll through AppInit_DLLs, so that processes that never touch ole32 (Excel)
 * still get the shim. Whichever instance comes first does the patching; the other stays passive. */
__declspec(dllexport) int ms365_shim_present = 1;
static BOOL other_instance_active(HINSTANCE inst)
{
    char path[MAX_PATH]; const char *base;
    if (!GetModuleFileNameA(inst, path, sizeof(path))) return FALSE;
    base = strrchr(path, '\\'); base = base ? base + 1 : path;
    const char *other = lstrcmpiA(base, "ms365shim.dll") == 0 ? "ole32.dll" : "ms365shim.dll";
    HMODULE h = GetModuleHandleA(other);
    return h && h != inst && GetProcAddress(h, "ms365_shim_present") != NULL;
}

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(inst);
        if (other_instance_active(inst)) { OutputDebugStringA("ms365 ole32 shim: second instance, passive"); return TRUE; }
        /* capture the genuine entry points before any import slot is rewritten */
        HMODULE k32 = GetModuleHandleA("kernel32.dll");
        real_GetProcAddress = (pGetProcAddress)GetProcAddress(k32, "GetProcAddress");
        real_GetModuleBaseNameA = (pGetModuleBaseNameA_t)GetProcAddress(k32, "K32GetModuleBaseNameA");
        real_GetDllDirectoryA = (pGetDllDirectoryA)GetProcAddress(k32, "GetDllDirectoryA");
        real_GetDllDirectoryW = (pGetDllDirectoryW)GetProcAddress(k32, "GetDllDirectoryW");
        GetEnvironmentVariableA("MS365_TRACE_MODULE", g_trace_mod_name, sizeof(g_trace_mod_name));
        HMODULE ntdll = GetModuleHandleA("ntdll.dll");
        pLdrRegisterDllNotification reg = ntdll ? (pLdrRegisterDllNotification)GetProcAddress(ntdll, "LdrRegisterDllNotification") : NULL;
        if (reg) reg(0, on_dll_notify, NULL, &notify_cookie);
        else OutputDebugStringA("ms365 ole32 shim: LdrRegisterDllNotification unavailable");
        patch_loaded_modules();
        for (size_t i = 0; i < sizeof(HIDE) / sizeof(HIDE[0]); i++) hide_export(HIDE[i].dll, HIDE[i].fn);
    }
    return TRUE;
}
