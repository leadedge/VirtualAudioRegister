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
//

#include <windows.h>
#include <Shlobj.h> // For IsUserAnAdmin
#include <Shlwapi.h> // For PathRemoveFileSpec
#include "resource.h" // Dialog layout
#include "SpoutGL\SpoutUtils.h" // Registry functions (local folder)

#pragma comment(lib, "Shlwapi.Lib")

using namespace spoututils;

static HINSTANCE g_hInst = nullptr; // Used for LoadImage
bool SetRegsvr32(char* dllpath, wchar_t* RegPath, bool use32Bit); // To register/Unregister

// Settings dialog
BOOL CALLBACK RegisterDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);


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


// Register of unregister the dll using regsvr32
bool SetRegsvr32(char * dllpath, wchar_t* wPath, bool use32Bit)
{
	PROCESS_INFORMATION pi ={};
	STARTUPINFO si ={sizeof(STARTUPINFO)};
	ZeroMemory((void*)&si, sizeof(STARTUPINFO));
	si.cb = sizeof(STARTUPINFO);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE; // hide the regsvr32 console window
	std::wstring cmdstring;
	bool bSuccess = false;

	// To register a 32-bit DLL
	//    use the 32-bit regsvr32 from SysWow64 (C:\Windows\SysWow64\regsvr32.exe)
	// To register a 64-bit DLL
	//    use the 64-bit regsvr32 from System32 (C:\Windows\System32\regsvr32.exe)

	// Choose regsvr32 path
    std::wstring regsvrPath = use32Bit
        ? L"C:\\Windows\\SysWow64\\regsvr32.exe"
        : L"C:\\Windows\\System32\\regsvr32.exe";

	if (dllpath && !wPath) {
		
		// Un-register
		wchar_t RegPath[MAX_PATH]={};
		mbstowcs(RegPath, dllpath, MAX_PATH);

		cmdstring = regsvrPath;
		cmdstring += L" /u ";

		cmdstring += L"\"";
		cmdstring += RegPath;
		cmdstring += L"\"";
	}
	else if(wPath && !dllpath) {

		// Register
		cmdstring = regsvrPath;
		cmdstring += L" ";

		cmdstring += L"\"";
		cmdstring += wPath;
		cmdstring += L"\"";
	}

	// Open regsvr32 and wait for completion
	HANDLE hProcess = NULL; // Handle from CreateProcess
	DWORD dwExitCode = 0; // Exit code to check completion

	BOOL success = CreateProcess(NULL, (LPWSTR)cmdstring.c_str(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
	if (!success) {
        DWORD err = GetLastError();
		SpoutMessageBox("CreateProcess failed", "Error 0x%7.7X", err);
        return false;
    }

	// Wait for completion
	WaitForSingleObject(pi.hProcess, INFINITE);

	DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    if (pi.hThread) CloseHandle(pi.hThread);
    if (pi.hProcess) CloseHandle(pi.hProcess);

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
	char str1[1024]={};
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
		return TRUE; // return TRUE unless you set the focus to a control

	case WM_COMMAND:

		switch (LOWORD(wParam)) {
	
			//
			// Register or un-register both 32 both and 64 bit versions
			//
			case IDC_CHECK:
			{
				if (!IsUserAnAdmin()) {
					MessageBoxA(NULL, "Administrator privileges are required\nRestart with RH click and 'Run as administrator'", "Warning", MB_OK);
					break;
				}

				// Get the new virtual audio device files
				wchar_t wpath[MAX_PATH]={};
				GetModuleFileName(NULL, wpath, MAX_PATH);
				PathRemoveFileSpec(wpath);

				wchar_t VirtualAudio32path[MAX_PATH]={};
				wchar_t VirtualAudio64path[MAX_PATH]={};
				wcscpy_s(VirtualAudio32path, MAX_PATH, wpath);
				wcscat_s(VirtualAudio32path, MAX_PATH, L"\\VirtualAudioDevice32\\audio_sniffer.dll");
				wcscpy_s(VirtualAudio64path, MAX_PATH, wpath);
				wcscat_s(VirtualAudio64path, MAX_PATH, L"\\VirtualAudioDevice64\\audio_sniffer-x64.dll");

				// Do the new files exist ?
				if (_waccess(VirtualAudio32path, 0) == -1
					|| _waccess(VirtualAudio64path, 0) == -1) {
					// Files do not exist
					MessageBoxA(NULL, "Virtual Audio Device files not found.\n\n\\VirtualAudioDevice32\\audio_sniffer.dll\n\\VirtualAudioDevice64\\audio_sniffer-x64.dll\n", "Warning", MB_OK);
					break;
				}

				//
				// Find paths for the files that have been previously registered from the registry
				//

				// Is 'virtual-audio-device' registered ?
				if (FindSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\\Classes\\CLSID\\{8E146464-DB61-4309-AFA1-3578E927E935}")
					|| FindSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}")) {
					// Is 32 bit registered
					char dllpath[MAX_PATH]={};
					if (ReadPathFromRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Classes\\CLSID\\{8E14549B-DB61-4309-AFA1-3578E927E935}\\InprocServer32", "", dllpath, MAX_PATH)) {
						// LJ DEBUG
						if(SetRegsvr32(dllpath, nullptr, true)) {
							SetDlgItemTextA(hDlg, IDC_CHECK, "Register");
						}
					}
#ifdef _WIN64
					// Repeat for 64 bit
					if (ReadPathFromRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\\Classes\\CLSID\\{8E146464-DB61-4309-AFA1-3578E927E935}\\InprocServer32", "", dllpath, MAX_PATH)) {
						// LJ DEBUG
						if (SetRegsvr32(dllpath, nullptr, false)) {
							SetDlgItemTextA(hDlg, IDC_CHECK, "Register");
						}
					}
#endif
				} // endif registered
				else {
					// 32 bit register
					// LJ DEBUG
					if (SetRegsvr32(nullptr, VirtualAudio32path, true)) {
						SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");
					}
#ifdef _WIN64
					// 64 bit register
					// LJ DEBUG
					if (SetRegsvr32(nullptr, VirtualAudio64path, true)) {
						SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");
					}
#endif
				} // endif not registered

			} // end case ID_CHECK
			break;

			case IDC_CHECK_HELP:
				sprintf_s(str1, 1024, "Register a virtual audio device that captures what you hear from the speakers.\n");
				strcat_s(str1, 1024, "The device is a DirectShow filter and can be used with FFmpeg to record the audio.\n");
				strcat_s(str1, 1024, "Developed by <a href=\"https://github.com/rdp/virtual-audio-capture-grabber-device\">Roger Pack</a>.\n\n");
				strcat_s(str1, 1024, "If 'virtual-audio-device' has not been registered, click the 'Register' button.\n");
				strcat_s(str1, 1024, "This will register both 32 bit and 64 bit versions.\n\n");

				strcat_s(str1, 1024, "You will see either confirmation of success or details of any error.\n");
				strcat_s(str1, 1024, "If an error occurs, the number is shown to help trace the problem.\n\n");
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


