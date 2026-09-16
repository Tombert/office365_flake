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
#include <windows.h>
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

typedef FARPROC (WINAPI *pGetProcAddress)(HMODULE, LPCSTR);
static pGetProcAddress real_GetProcAddress;
typedef DWORD (WINAPI *pGetModuleBaseNameA_t)(HANDLE, HMODULE, LPSTR, DWORD);
static pGetModuleBaseNameA_t real_GetModuleBaseNameA;

static const struct patch *wrapped_lookup(HMODULE h, LPCSTR name);
static FARPROC WINAPI my_GetProcAddress(HMODULE h, LPCSTR name)
{
    FARPROC p = real_GetProcAddress(h, name);
    if (name && ((ULONG_PTR)name >> 16) != 0) {
        /* functions we wrap (see PATCHES) must be wrapped for dynamic lookups too */
        const struct patch *w = p ? wrapped_lookup(h, name) : NULL;
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

static const struct patch PATCHES[] = {
    { "kernel32.dll", "GetProcAddress",              (void *)my_GetProcAddress },
    { "winhttp.dll",  "WinHttpSetOption",            (void *)my_WinHttpSetOption },
    { "winhttp.dll",  "WinHttpQueryOption",          (void *)my_WinHttpQueryOption },
    { "wininet.dll",  "InternetSetOptionW",          (void *)my_InternetSetOptionW },
    { "wininet.dll",  "InternetSetOptionA",          (void *)my_InternetSetOptionA },
    { "wininet.dll",  "InternetQueryOptionW",        (void *)my_InternetQueryOptionW },
    { "wininet.dll",  "InternetQueryOptionA",        (void *)my_InternetQueryOptionA },
    { "kernel32.dll", "SetFileShortNameW",           (void *)my_SetFileShortNameW },
    { "kernel32.dll", "SetFileShortNameA",           (void *)my_SetFileShortNameA },
    { "kernel32.dll", "FindPackagesByPackageFamily", (void *)my_FindPackagesByPackageFamily },
    { "kernel32.dll", "SetThreadpoolTimerEx",        (void *)my_SetThreadpoolTimerEx },
};
#define NPATCHES (sizeof(PATCHES) / sizeof(PATCHES[0]))

/* PATCHES entries that wrap an existing export (not GetProcAddress itself), keyed by module.
 * No API calls in here: Office's App-V layer (AppVIsvSubsystems64) detours kernel32 entry points
 * such as GetModuleHandleA and fails fast when they are re-entered from inside a GetProcAddress
 * that it is making itself. Module handles are recorded by the loader notification instead. */
static struct wrapmod { const char *dll; HMODULE h; } WRAPMODS[] = {
    { "winhttp.dll", NULL },
    { "wininet.dll", NULL },
};
#define NWRAPMODS (sizeof(WRAPMODS) / sizeof(WRAPMODS[0]))
static const struct patch *wrapped_lookup(HMODULE h, LPCSTR name)
{
    for (size_t m = 0; m < NWRAPMODS; m++) {
        if (!WRAPMODS[m].h || WRAPMODS[m].h != h) continue;
        for (size_t i = 0; i < NPATCHES; i++)
            if (PATCHES[i].dll == WRAPMODS[m].dll || (lstrcmpiA(PATCHES[i].dll, WRAPMODS[m].dll) == 0 && strcmp(PATCHES[i].fn, name) == 0)) {
                if (strcmp(PATCHES[i].fn, name) == 0) return &PATCHES[i];
            }
        return NULL;
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
        } else {
            real_InternetSetOptionW = (pInternetSetOption)real_GetProcAddress(h, "InternetSetOptionW");
            real_InternetSetOptionA = (pInternetSetOption)real_GetProcAddress(h, "InternetSetOptionA");
            real_InternetQueryOptionW = (pInternetQueryOption)real_GetProcAddress(h, "InternetQueryOptionW");
            real_InternetQueryOptionA = (pInternetQueryOption)real_GetProcAddress(h, "InternetQueryOptionA");
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
            if (IMAGE_SNAP_BY_ORDINAL(ilt->u1.Ordinal)) continue;
            IMAGE_IMPORT_BY_NAME *ibn = (IMAGE_IMPORT_BY_NAME *)(base + ilt->u1.AddressOfData);
            const char *fn = (const char *)ibn->Name;
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

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(inst);
        /* capture the genuine entry points before any import slot is rewritten */
        HMODULE k32 = GetModuleHandleA("kernel32.dll");
        real_GetProcAddress = (pGetProcAddress)GetProcAddress(k32, "GetProcAddress");
        real_GetModuleBaseNameA = (pGetModuleBaseNameA_t)GetProcAddress(k32, "K32GetModuleBaseNameA");
        HMODULE ntdll = GetModuleHandleA("ntdll.dll");
        pLdrRegisterDllNotification reg = ntdll ? (pLdrRegisterDllNotification)GetProcAddress(ntdll, "LdrRegisterDllNotification") : NULL;
        if (reg) reg(0, on_dll_notify, NULL, &notify_cookie);
        else OutputDebugStringA("ms365 ole32 shim: LdrRegisterDllNotification unavailable");
        patch_loaded_modules();
        for (size_t i = 0; i < sizeof(HIDE) / sizeof(HIDE[0]); i++) hide_export(HIDE[i].dll, HIDE[i].fn);
    }
    return TRUE;
}
