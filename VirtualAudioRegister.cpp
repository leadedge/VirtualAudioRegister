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
//	versions of the dll as DirectShow Filter ".ax" files.
//
//
//		30.03.23 - Started based on SpoutCamSettings
//				   Version 1.000
//

#include <windows.h>
#include <Shlobj.h> // For IsUserAnAdmin
#include "resource.h" // Dialog layout
#include "SpoutGL\SpoutUtils.h" // Registry functions (local folder)
using namespace spoututils;

static HINSTANCE g_hInst = nullptr; // Used for LoadImage
bool SetRegistry(char* dllpath, wchar_t* RegPath); // Tor register/Unregister

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


// Register of unregister the dll using regsvr32 silent
bool SetRegistry(char * dllpath, wchar_t* wPath)
{
	PROCESS_INFORMATION pi ={};
	STARTUPINFO si ={sizeof(STARTUPINFO)};
	ZeroMemory((void*)&si, sizeof(STARTUPINFO));
	si.cb = sizeof(STARTUPINFO);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE; // hide the regsvr32 console window
	std::wstring cmdstring;
	bool bSuccess = false;

	if (dllpath && !wPath) {
		// Un-register
		wchar_t RegPath[MAX_PATH]={};
		mbstowcs(RegPath, dllpath, MAX_PATH);
		cmdstring = L"regsvr32.exe /u /s ";
		cmdstring += L"\"";
		cmdstring += RegPath;
		cmdstring += L"\"";
	}
	else if(wPath && !dllpath) {
		// Register
		cmdstring = L"regsvr32.exe /s ";
		cmdstring += L"\"";
		cmdstring += wPath;
		cmdstring += L"\"";
	}

	// Open regsvr32 and wait for completion
	HANDLE hProcess = NULL; // Handle from CreateProcess
	DWORD dwExitCode = 0; // Exit code to check completion

	if (CreateProcess(NULL, (LPWSTR)cmdstring.c_str(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
		hProcess = pi.hProcess;
		if (hProcess) {
			do {
				GetExitCodeProcess(hProcess, &dwExitCode);
			} while (dwExitCode == STILL_ACTIVE);
			bSuccess = true;
			CloseHandle(pi.hProcess);
		}
		if (pi.hThread) CloseHandle(pi.hThread);
	}

	return bSuccess;

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
			// Change title on IDC_CHECK button
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
				wcscat_s(VirtualAudio32path, MAX_PATH, L"\\VirtualAudioDevice32\\audio_sniffer.ax");
				wcscpy_s(VirtualAudio64path, MAX_PATH, wpath);
				wcscat_s(VirtualAudio64path, MAX_PATH, L"\\VirtualAudioDevice64\\audio_sniffer-x64.ax");

				// Do the new files exist ?
				if (_waccess(VirtualAudio32path, 0) == -1
					|| _waccess(VirtualAudio64path, 0) == -1) {
					// Files do not exist
					MessageBoxA(NULL, "Virtual Audio Device files not found.\n\n\\VirtualAudioDevice32\\audio_sniffer.ax\n\\VirtualAudioDevice64\\audio_sniffer-x64.ax\n", "Warning", MB_OK);
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
						if(SetRegistry(dllpath, nullptr)) {
							SetDlgItemTextA(hDlg, IDC_CHECK, "Register");
							MessageBoxA(NULL, "virtual-audio-device 32 bit un-registered successfully.\n", "Information", MB_OK);
						}
						else {
							MessageBoxA(NULL, "Problem un-registering 32 bit virtual-audio-device\nClose any programs that could be using it\nand try again.", "Warning", MB_OK);
						}
					}
#ifdef _WIN64
					// Repeat for 64 bit
					if (ReadPathFromRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\\Classes\\CLSID\\{8E146464-DB61-4309-AFA1-3578E927E935}\\InprocServer32", "", dllpath, MAX_PATH)) {
						if (SetRegistry(dllpath, nullptr)) {
							SetDlgItemTextA(hDlg, IDC_CHECK, "Register");
							MessageBoxA(NULL, "virtual-audio-device 64 bit un-registered successfully.\n", "Information", MB_OK);
						}
						else {
							MessageBoxA(NULL, "Problem un-registering 64 bit virtual-audio-device\nClose any programs that could be using it\nand try again.", "Warning", MB_OK);
						}
					}
#endif
				} // endif registered
				else {

					// 32 bit register
					if (SetRegistry(nullptr, VirtualAudio32path)) {
						SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");
						MessageBoxA(NULL, "virtual-audio-device 32 bit registered successfully.\n", "Information", MB_OK);
					}
					else {
						MessageBoxA(NULL, "Problem registering 32 bit virtual-audio-device\nClose any programs that could be using it\nand try again.", "Warning", MB_OK);
					}
#ifdef _WIN64
					// 64 bit register
					if (SetRegistry(nullptr, VirtualAudio64path)) {
						SetDlgItemTextA(hDlg, IDC_CHECK, "UnRegister");
						MessageBoxA(NULL, "virtual-audio-device 64 bit registered successfully.\n", "Information", MB_OK);
					}
					else {
						MessageBoxA(NULL, "Problem registering 64 bit virtual-audio-device\nClose any programs that could be using it\nand try again.", "Warning", MB_OK);
					}
#endif
				} // endif not registered

			} // end case ID_CHECK
			break;

			case IDC_CHECK_HELP:
				sprintf_s(str1, 1024, "Register a virtual audio device that captures what you hear from the speakers. ");
				strcat_s(str1, 1024, "The device is a DirectShow filter and can be used with FFmpeg to record the audio. ");
				strcat_s(str1, 1024, "Developed by Roger Pack :\n\n");
				strcat_s(str1, 1024, "https://github.com/rdp/virtual-audio-capture-grabber-device\n\n");
				strcat_s(str1, 1024, "If 'virtual-audio-device' has not been registered, click the 'Register' button. ");
				strcat_s(str1, 1024, "This will register both 32 bit and 64 bit versions.\n\n");
				strcat_s(str1, 1024, "After confirmation of success, the button will show 'UnRegister'. ");
				strcat_s(str1, 1024, "Click 'Unregister' to remove virtual-audio-device from the system.\n\n");
				strcat_s(str1, 1024, "To update to new copies of the virtual-audio-device dll, click 'Unregister' and then 'Register' again.\n\n");
				strcat_s(str1, 1024, "https://spout.zeal.co/\n\n");
				MessageBoxA(hDlg, str1, "Registration", MB_OK | MB_ICONINFORMATION);
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


