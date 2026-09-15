/* Implementations for entry points the builtin ole32/combase of this Wine build lack. */
#include <windows.h>

/* Office (mso30win32client) registers an activation filter at startup and dereferences the result
 * of GetProcAddress unconditionally. Filters only matter for app-container activation; accept it. */
__declspec(dllexport) HRESULT WINAPI CoRegisterActivationFilter(void *filter)
{
    (void)filter;
    return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)inst; (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(inst);
    return TRUE;
}
