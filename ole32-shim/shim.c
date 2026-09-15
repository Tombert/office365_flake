/*
 * ole32.dll shim for Microsoft Office under Wine/Proton.
 *
 * 1. Exports: every export of the builtin ole32 is forwarded (see gen-def.py) to combase or to
 *    ole32_wine.dll (a renamed copy of the builtin), plus entry points the builtin lacks.
 * 2. Stub patching: Wine marks unimplemented functions as stubs that raise a non-continuable
 *    exception, which kills Office. Since ole32 is loaded early by every Office process, DllMain
 *    overwrites the first bytes of each listed stub with a jump to a replacement that behaves the
 *    way the real API does for an unprivileged caller.
 *
 * x86_64 only. Add entries to PATCHES as new "unimplemented function X.Y called" aborts turn up.
 */
#include <windows.h>
#include <string.h>

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

struct patch { const char *dll; const char *fn; void *repl; };
static const struct patch PATCHES[] = {
    { "kernel32.dll", "SetFileShortNameW", (void *)my_SetFileShortNameW },
    { "kernel32.dll", "SetFileShortNameA", (void *)my_SetFileShortNameA },
};

static void apply(const struct patch *p)
{
    HMODULE h = GetModuleHandleA(p->dll);
    if (!h) return; /* only patch what is already loaded; never LoadLibrary under the loader lock */
    BYTE *fn = (BYTE *)GetProcAddress(h, p->fn);
    if (!fn) return;
    if (fn[0] == 0x48 && fn[1] == 0xB8 && fn[10] == 0xFF && fn[11] == 0xE0) return; /* done before */
    DWORD old;
    if (!VirtualProtect(fn, 16, PAGE_EXECUTE_READWRITE, &old)) return;
    fn[0] = 0x48; fn[1] = 0xB8;                    /* mov rax, imm64 */
    memcpy(fn + 2, &p->repl, sizeof(void *));
    fn[10] = 0xFF; fn[11] = 0xE0;                  /* jmp rax */
    VirtualProtect(fn, 16, old, &old);
    FlushInstructionCache(GetCurrentProcess(), fn, 16);
}

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(inst);
        for (size_t i = 0; i < sizeof(PATCHES) / sizeof(PATCHES[0]); i++) apply(&PATCHES[i]);
    }
    return TRUE;
}
