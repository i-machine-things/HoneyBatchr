#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <wchar.h>

/* GUI launcher — no console window.
   Resolves paths relative to its own location so the install can live anywhere. */
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                    LPWSTR lpCmdLine, int nCmdShow) {
    wchar_t base[MAX_PATH];
    GetModuleFileNameW(NULL, base, MAX_PATH);
    wchar_t *last = wcsrchr(base, L'\\');
    if (last) *last = L'\0';

    /* Build: "<base>\runtime\python.exe" "<base>\app\main.py" */
    wchar_t cmd[MAX_PATH * 3];
    wcscpy(cmd, L"\"");
    wcscat(cmd, base);
    wcscat(cmd, L"\\runtime\\python.exe\" \"");
    wcscat(cmd, base);
    wcscat(cmd, L"\\app\\main.py\"");

    STARTUPINFOW        si = {sizeof(si)};
    PROCESS_INFORMATION pi = {0};
    si.dwFlags     = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    CreateProcessW(NULL, cmd, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
    if (pi.hProcess) CloseHandle(pi.hProcess);
    if (pi.hThread)  CloseHandle(pi.hThread);
    return 0;
}
