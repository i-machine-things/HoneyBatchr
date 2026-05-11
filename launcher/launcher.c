/*
 * HoneyBatchr launcher
 *
 * Finds runtime\pythonw.exe relative to this executable and launches
 * app\main.py with no visible console window. The working directory is
 * set to the install root so relative paths inside the app resolve correctly.
 *
 * Build (MinGW, GitHub Actions windows-latest):
 *   windres launcher/launcher.rc -O coff -o launcher_res.o
 *   gcc -mwindows -O2 -o HoneyBatchr.exe launcher/launcher.c launcher_res.o
 */

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <wchar.h>

int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPSTR lpCmd, int nShow)
{
    wchar_t exe_path[MAX_PATH];
    GetModuleFileNameW(NULL, exe_path, MAX_PATH);
    wchar_t *last_sep = wcsrchr(exe_path, L'\\');
    if (last_sep) *last_sep = L'\0';

    wchar_t python[MAX_PATH];
    wchar_t script[MAX_PATH];
    _snwprintf_s(python, MAX_PATH, _TRUNCATE,
                 L"%s\\runtime\\pythonw.exe", exe_path);
    _snwprintf_s(script, MAX_PATH, _TRUNCATE,
                 L"%s\\app\\main.py", exe_path);

    wchar_t cmdline[MAX_PATH * 2];
    _snwprintf_s(cmdline, MAX_PATH * 2, _TRUNCATE,
                 L"\"%s\" \"%s\"", python, script);

    STARTUPINFOW si;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);

    PROCESS_INFORMATION pi;
    ZeroMemory(&pi, sizeof(pi));

    BOOL created = CreateProcessW(
        python,
        cmdline,
        NULL, NULL, FALSE, 0, NULL,
        exe_path,
        &si, &pi
    );

    if (!created) {
        DWORD err = GetLastError();
        wchar_t msg[512];
        _snwprintf_s(msg, 512, _TRUNCATE,
                     L"Failed to start HoneyBatchr.\n\n"
                     L"Python: %s\n"
                     L"Script: %s\n"
                     L"Error code: %lu",
                     python, script, (unsigned long)err);
        MessageBoxW(NULL, msg, L"HoneyBatchr Launch Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return 0;
}
