//
//					VirtualAudioRegister
//
//	A dialog application to register 'virtual-audio-device',
//	a DirectShow filter to capture system audio, developed by Roger Pack.
//
//	https://github.com/rdp/virtual-audio-capture-grabber-device/issues/38#issuecomment-1489671448
//
//	This application should be built x64 to register both versions directly using the registry.
//	A 64 bit Windows system is required. _WIN64 is checked at compile, not at runtime
//	If built with Visual Studio "C++ > Code Generation > Runtime Library /MT",
//	the Visual Studio runtime dlls are not necessary.
//	The application has admin privileges and UAC could intercept but allow.
//
//	The application has to be located together with folders 
//	"VirtualAudioDevice32" and "VirtualAudioDevice64" with 32 bit and 64 bit
//	versions of the DirectShow Filter dll files.
//
//
//		30.03.23 - Started based on SpoutCamSettings
//				   Version 1.000
//		12.08.23 - Changed to requireAdministrator privileges
//		05.09.23 - Configuration all - requireAdministrator privileges
//				   Version 1.001
//		22.04.24 - Update SpoutUtils
//				   Include Shlwapi, removed from SpoutUtils
//				   Increase font size from 8 to 9
//				   Replace MessageBox with SpoutMessageBox for help dialog
//				   Version 1.002
//		10.09.24 - Update SpoutUtils 2.007.014 
//				   Correct ReadPathFromRegistry
//				   "valuename" argument can be null for the (Default) key string
//		11.09.24 - Change .ax to .dll for filter files
//				   Update resource file version number
//				   Version 1.003
//		15.04.25 - Use 32bit or 64bit Regsrvr32 for 32 or 64 bit register
//				   Use Regsrvr32 without -s (silent) option so that all messages are displayed.
//				   Show error number on failure for diagnostics.
//				   Remove successful/unsuccessful messageboxes.
//				   Use WaitForSingleObject to wait for process completion.
//				   Correct url to Roger Pack repo for virtual-audio-capture-grabber-device
//				   Version 1.004
//		16.04.25 - Add "Include 32 bit" checkbox to register for both 32 and 64 bit programs
//				   Change logic for Register/UnRegister depending on button state
//				   Restore silent option for regsrv32
//				   Change font from "Ms Shell Dlg" to "Segoe UI"
//				   Version 1.005
//

#include <windows.h>
#include <Shlobj.h> // For IsUserAnAdmin
#include <Shlwapi.h> // For PathRemoveFileSpec
#include "resource.h" // Dialog layout
#include "SpoutGL\SpoutUtils.h" // Registry functions (local folder)

#pragma comment(lib, "Shlwapi.Lib")

using namespace spoututils;

 // register/unregister 32 or 64 bit
bool SetRegsvr32(char* dllpath, wchar_t* RegPath, bool b32bit);

// Settings dialog
BOOL CALLBACK RegisterDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);

static bool systemType = false; // true - register for both 64bit and 32bit programs, false 64bit only
static HINSTANCE g_hInst = nullptr; // Used for LoadImage


int WINAPI WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nCmdShow)
{
	g_hInst = hInstance;

	// Console for debugging
	// OpenSpoutConsole();

	// This is a modal dialog and will stop this application at this point
	if (DialogBoxParamA(hInstance, MAKEINTRESOURCEA(IDD_AUDIOBOX), 0, (DLGPROC)RegisterDialogProc, 0)) {
		return 1;
	}

	return 0;
}


// Register of unregister a 32bit or 64bit dll using regsvr32
bool SetRegsvr32(char * dllpath, wchar_t* wPath, bool b32bit)
{
	PROCESS_INFORMATION pi{};
	STARTUPINFO si ={sizeof(STARTUPINFO)};
	ZeroMemory((void*)&si, sizeof(STARTUPINFO));
	si.cb = sizeof(STARTUPINFO);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE; // hide the regsvr32 console window
	std::wstring cmdstring;
	std::wstring regsvrPath;
	bool bSuccess = false;

	// To register a 32-bit DLL
	//    use the 32-bit regsvr32 from SysWow64 (C:\Windows\SysWow64\regsvr32.exe)
	// To register a 64-bit DLL
	//    use the 64-bit regsvr32 from System32 (C:\Windows\System32\regsvr32.exe)
	wchar_t windowsDir[MAX_PATH]{};
	if (b32bit)
		GetSystemWow64DirectoryW(windowsDir, MAX_PATH);
	else
		GetSystemDirectoryW(windowsDir, MAX_PATH);
	regsvrPath = std::wstring(windowsDir);
	regsvrPath += L"\\regsvr32.exe";

	if (dllpath && !wPath) {
		// Un-register
		wchar_t RegPath[MAX_PATH]{};
		mbstowcs(RegPath, dllpath, MAX_PATH);
		cmdstring = regsvrPath;
		cmdstring += L" /u /s ";
		cmdstring += L"\"";
		cmdstring += RegPath;
		cmdstring += L"\"";
	}
	else if(wPath && !dllpath) {
		// Register
		cmdstring = regsvrPath;
		cmdstring += L" /s ";
		cmdstring += L"\"";
		cmdstring += wPath;
		cmdstring += L"\"";
	}

	// Open regsvr32 and wait for completion
	HANDLE hProcess = NULL; // Handle from CreateProcess
	DWORD dwExitCode = 0; // Exit code to check completion

	BOOL success = CreateProcess(NULL, (LPWSTR)cmdstring.c_str(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
	if (!success || !pi.hProcess) {
        DWORD err = GetLastError();
		SpoutMessageBox("CreateProcess failed", "Error 0x%7.7X", err);
        return false;
    }

	// Wait for completion
	WaitForSingleObject(pi.hProcess, INFINITE);

	DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hProcess);

	if (pi.hThread) CloseHandle(pi.hThread);
	if (exitCode != 0) {
		SpoutMessageBox("Regsvr32 failed", "Exit code 0x%7.7X", exitCode);
        return false;
    }

	return true;

}


// Message handler for dialog
BOOL CALLBACK RegisterDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam); // suppress warning
	char str1[1024]{};
	HICON hIcon1 = NULL;
	HBRUSH hbrBkgnd = NULL;

	switch (message) {

		// Owner draw button
		case WM_CTLCOLORSTATIC:
		{
			HDC hdcStatic = (HDC)wParam;
			// Make Version info light grey
			if (GetDlgItem(hDlg, IDC_VERS) == (HWND)lParam) {
				SetTextColor(hdcStatic, RGB(96, 96, 96));
				SetBkColor(hdcStatic, RGB(240, 240, 240));
				if (hbrBkgnd == NULL)
					hbrBkgnd = CreateSolidBrush(RGB(240, 240, 240));
				return (INT_PTR)hbrBkgnd;
			}
		}
		break;

	case WM_INITDIALOG:

		// Set the window icon
		// http://blog.barthe.ph/2009/07/17/wmseticon/
		hIcon1 = (HICON)LoadImage(g_hInst,
								MAKEINTRESOURCE(IDI_SPOUTICON),
								IMAGE_ICON,
								GetSystemMetrics(SM_CXICON),
								GetSystemMetrics(SM_CXICON),
								0);
		if (hIcon1) {
			SendMessage(hDlg, WM_SETICON, ICON_BIG, (LPARAM)hIcon1);
			SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon1);
		}

		HICON hIcon2;
		hIcon2 = (HICON)LoadImage(g_hInst,
								MAKEINTRESOURCE(IDI_SPOUTICON),
								IMAGE_ICON,
								GetSystemMetrics(SM_CXICON),
								GetSystemMetrics(SM_CXICON),
								0);
		if(hIcon2)
			SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon2);

		// Is 'virtual-audio-device' registered ?
		if (FindSubKey(HKEY_LOCAL_MACHINE, "\\SOFTWARE\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}")
		 || FindSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}")) {
			// Change title on IDC_CHECK button (default in resource.rc is "Register")
			SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");
		}

		// 32 bit checkbox
		systemType = false;
		if (FindSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}")) {
			systemType = true;
		}
		CheckDlgButton(hDlg, IDC_32, systemType); // true - both 32 & 64 bit

		return TRUE; // return TRUE unless you set the focus to a control

	case WM_COMMAND:

		switch (LOWORD(wParam)) {

			//
			// System type checkbox
			// True - register for both 64bit and 32bit programs
			case IDC_32:
				if(IsDlgButtonChecked(hDlg, IDC_32))
					systemType = true;
				else
					systemType = false;
				break;
	
			//
			// Register or un-register both 32 both and 64 bit versions
			//
			case IDC_CHECK: {

				if (!IsUserAnAdmin()) {
					MessageBoxA(NULL, "Administrator privileges are required\nRestart with RH click and 'Run as administrator'", "Warning", MB_OK);
					break;
				}

				// Get the new virtual audio device files
				wchar_t wpath[MAX_PATH]{};
				GetModuleFileName(NULL, wpath, MAX_PATH);
				PathRemoveFileSpec(wpath);

				wchar_t VirtualAudio32path[MAX_PATH]{};
				wchar_t VirtualAudio64path[MAX_PATH]{};
				wcscpy_s(VirtualAudio32path, MAX_PATH, wpath);
				wcscat_s(VirtualAudio32path, MAX_PATH, L"\\VirtualAudioDevice32\\audio_sniffer.dll");
				wcscpy_s(VirtualAudio64path, MAX_PATH, wpath);
				wcscat_s(VirtualAudio64path, MAX_PATH, L"\\VirtualAudioDevice64\\audio_sniffer-x64.dll");

				// Do the new files exist ?
				if (_waccess(VirtualAudio32path, 0) == -1 || _waccess(VirtualAudio64path, 0) == -1) {
					// Files do not exist
					MessageBoxA(NULL, "Virtual Audio Device files not found.\n\n\\VirtualAudioDevice32\\audio_sniffer.dll\n\\VirtualAudioDevice64\\audio_sniffer-x64.dll\n", "Warning", MB_OK);
					break;
				}

				// Is the button "Register" or "UnRegister"
				GetDlgItemTextA(hDlg, IDC_CHECK, str1, MAX_PATH);

				if (strcmp(str1, "UnRegister") == 0) {

					//
					// Always un-register both 32 and 64 bit
					//

					// Find paths for the files that have been previously registered from the registry
					char dllpath[MAX_PATH] {};
					// 32 bit
					if (FindSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}")) {
						// Un-register 32
						if (ReadPathFromRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}\\InprocServer32", "", dllpath, MAX_PATH)) {
							if (SetRegsvr32(dllpath, nullptr, true)) { // 32 bit register
								SpoutMessageBox(NULL, "32 bit un-registered successfuly", "VirtualAudioRegister", MB_ICONINFORMATION | MB_OK);
							} else {
								int err = GetLastError();
								SpoutMessageBox("Warning", "Error for 32 bit un-register %d (0x%6.6X)", err, err);
							}
						}
					}

					//
					// Repeat for 64 bit
					//
					if (ReadPathFromRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\\Classes\\CLSID\\{8E146464-DB61-4309-AFA1-3578E927E935}\\InprocServer32", "", dllpath, MAX_PATH)) {
						// Un-register 64
						if (SetRegsvr32(dllpath, nullptr, false)) { // 64 bit unregister
							SpoutMessageBox(NULL, "64 bit un-registered successfuly", "VirtualAudioRegister", MB_ICONINFORMATION | MB_OK);
						} else {
							int err = GetLastError();
							SpoutMessageBox("Warning", "Error for 64 bit un-register %d (0x%6.6X)", err, err);
						}
					}
					SetDlgItemTextA(hDlg, IDC_CHECK, "Register");

				} // Endif button was "UnRegister"
				else {

					//
					// Register if un-registered
					//

					// Register 32 bit if selected
					if (systemType) {
						if (SetRegsvr32(nullptr, VirtualAudio32path, true)) { // 32 bit register
							SpoutMessageBox(NULL, "32 bit registered successfuly", "VirtualAudioRegister", MB_ICONINFORMATION | MB_OK);
						}
						else {
							int err = GetLastError();
							SpoutMessageBox("Warning", "Error for 32 bit register %d (0x%6.6X)", err, err);
						}
					}

					// 64 bit
					if (SetRegsvr32(nullptr, VirtualAudio64path, false)) { // 64 bit register
						SpoutMessageBox(NULL, "64 bit registered successfuly", "VirtualAudioRegister", MB_ICONINFORMATION | MB_OK);
					} else {
						int err = GetLastError();
						SpoutMessageBox("Warning", "Error for 64 bit register %d (0x%6.6X)", err, err);
					}
					SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");

				} // Endif button was "Register"

			}  // end case ID_CHECK

			break;

			case IDC_CHECK_HELP:
				sprintf_s(str1, 1024, "Register a virtual audio device that captures what you hear from the speakers.\n");
				strcat_s(str1, 1024, "The device is a DirectShow filter and can be used with FFmpeg to record the audio.\n");
				strcat_s(str1, 1024, "Developed by <a href=\"https://github.com/rdp/virtual-audio-capture-grabber-device\">Roger Pack</a>.\n\n");

				strcat_s(str1, 1024, "Check \"Include 32 bit\" to register for use with both 32 bit and 64 bit programs.\n\n");

				strcat_s(str1, 1024, "If 'virtual-audio-device' has not been registered, click the 'Register' button.\n");
				strcat_s(str1, 1024, "You will see either confirmation of success or details of any error.\n");
				strcat_s(str1, 1024, "On success, the button will show 'UnRegister' which can then be used\n");
				strcat_s(str1, 1024, "to remove virtual-audio-device from the system.\n\n");

				strcat_s(str1, 1024, "To update to new copies of the virtual-audio-device dll,\nclick 'Unregister' and then 'Register' again.\n\n");
				strcat_s(str1, 1024, "                                            <a href=\"http://spout.zeal.co\">http://spout.zeal.co</a>\n\n");

				SpoutMessageBoxIcon(LoadIconA(GetModuleHandle(NULL), MAKEINTRESOURCEA(IDI_SPOUTICON)));
				SpoutMessageBox(hDlg, str1, "Registration", MB_OK | MB_USERICON);

				break;

			case IDOK:
				EndDialog(hDlg, LOWORD(wParam));
				return TRUE;

			default:
				return FALSE;
		}
		break;

	case WM_SYSCOMMAND:
		if (wParam == SC_CLOSE)	{
			EndDialog (hDlg, TRUE);
			return TRUE;
		}
		break;

	}

	return FALSE;

}

// The end..


