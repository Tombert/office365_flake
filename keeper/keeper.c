/*
 * ms365keeper.exe: holds a prefix's shared Wine session open while any Office app runs in it.
 *
 * The first app on a prefix starts the session, and umu's launcher service in that session's
 * container is how every later app and tool gets into the same wineserver. The service lives
 * exactly as long as the command it was started with, and when it stops it kills everything it
 * launched. So the first launch runs this instead of the app itself: it starts the app, then waits
 * until Wine has no user processes left (the event Wine's own services wait on) and exits, ending
 * the session only after the last app has closed, whichever one that is.
 *
 * Usage: ms365keeper.exe <command line of the app to start>
 */
#include <windows.h>

#define ProcessBasicInformation      0
#define ProcessWineMakeProcessSystem 1000   /* Wine's include/winternl.h */

/* PROCESS_BASIC_INFORMATION with the field names mingw's winternl.h leaves out */
struct basic_info {
    LONG      ExitStatus;
    void     *PebBaseAddress;
    ULONG_PTR AffinityMask;
    LONG      BasePriority;
    ULONG_PTR UniqueProcessId;
    ULONG_PTR InheritedFromUniqueProcessId;
};

LONG NTAPI NtSetInformationProcess(HANDLE process, ULONG class, void *info, ULONG size);
LONG NTAPI NtQueryInformationProcess(HANDLE process, ULONG class, void *info, ULONG size, ULONG *ret);

/* Stop counting <process> as a user process; returns the event that is set when no user
 * processes are left, or NULL. */
static HANDLE make_process_system(HANDLE process)
{
    HANDLE event = NULL;
    if (NtSetInformationProcess(process, ProcessWineMakeProcessSystem, &event, sizeof(HANDLE *))) return NULL;
    return event;
}

/* the command line after argv[0], which (unlike the other arguments) has no escapes */
static WCHAR *skip_argv0(WCHAR *cmd)
{
    if (*cmd == '"') {
        for (cmd++; *cmd && *cmd != '"'; cmd++) ;
        if (*cmd) cmd++;
    } else {
        while (*cmd && *cmd != ' ' && *cmd != '\t') cmd++;
    }
    while (*cmd == ' ' || *cmd == '\t') cmd++;
    return cmd;
}

int WINAPI WinMain(HINSTANCE inst, HINSTANCE prev, LPSTR cmdline, int show)
{
    WCHAR *app = skip_argv0(GetCommandLineW());
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    struct basic_info info;
    HANDLE parent, shutdown;

    if (!*app) return 1;
    if (!CreateProcessW(NULL, app, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) return GetLastError();
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);

    /* Proton's umu.exe started us and waits for us; as a user process it would keep the session
     * from ever seeing its last user process go. */
    if (!NtQueryInformationProcess(GetCurrentProcess(), ProcessBasicInformation, &info, sizeof(info), NULL) &&
        (parent = OpenProcess(PROCESS_SET_INFORMATION, FALSE, (DWORD)info.InheritedFromUniqueProcessId))) {
        HANDLE event = make_process_system(parent);
        if (event) CloseHandle(event);
        CloseHandle(parent);
    }

    if (!(shutdown = make_process_system(GetCurrentProcess()))) return 1;
    WaitForSingleObject(shutdown, INFINITE);
    return 0;
}
