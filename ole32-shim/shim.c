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

/* Exported stubs to overwrite in place (12-byte "mov rax, imm64; jmp rax") when their DLL loads.
 * Unlike PATCHES these exist in the export table, but calling them aborts the process. */
static const struct patch STUBPATCHES[] = {
    { "oleacc.dll", "CreateStdAccessibleProxyW", (void *)my_CreateStdAccessibleProxyW },
    { "oleacc.dll", "CreateStdAccessibleProxyA", (void *)my_CreateStdAccessibleProxyA },
};
#define NSTUBPATCHES (sizeof(STUBPATCHES) / sizeof(STUBPATCHES[0]))

static const struct patch PATCHES[] = {
    { "kernel32.dll", "SetFileShortNameW",           (void *)my_SetFileShortNameW },
    { "kernel32.dll", "SetFileShortNameA",           (void *)my_SetFileShortNameA },
    { "kernel32.dll", "FindPackagesByPackageFamily", (void *)my_FindPackagesByPackageFamily },
    { "kernel32.dll", "SetThreadpoolTimerEx",        (void *)my_SetThreadpoolTimerEx },
};
#define NPATCHES (sizeof(PATCHES) / sizeof(PATCHES[0]))

/* ---- machinery --------------------------------------------------------------------------- */

static void dbg(const char *what, const char *mod, const char *fn)
{
    char buf[320];
    wsprintfA(buf, "ms365 ole32 shim: %s %s!%s", what, mod, fn);
    OutputDebugStringA(buf);
}

static int ieq(const char *a, const char *b) { return lstrcmpiA(a, b) == 0; }

/* Rewrite import slots of one module. */
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
    if (reason != 1 /* LDR_DLL_NOTIFICATION_REASON_LOADED */ || !data || !data->DllBase) return;
    char name[128] = "?";
    if (data->BaseDllName && data->BaseDllName->Buffer)
        WideCharToMultiByte(CP_ACP, 0, data->BaseDllName->Buffer, data->BaseDllName->Length / 2, name, sizeof(name) - 1, NULL, NULL);
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
        HMODULE ntdll = GetModuleHandleA("ntdll.dll");
        pLdrRegisterDllNotification reg = ntdll ? (pLdrRegisterDllNotification)GetProcAddress(ntdll, "LdrRegisterDllNotification") : NULL;
        if (reg) reg(0, on_dll_notify, NULL, &notify_cookie);
        else OutputDebugStringA("ms365 ole32 shim: LdrRegisterDllNotification unavailable");
        patch_loaded_modules();
        for (size_t i = 0; i < sizeof(HIDE) / sizeof(HIDE[0]); i++) hide_export(HIDE[i].dll, HIDE[i].fn);
    }
    return TRUE;
}
