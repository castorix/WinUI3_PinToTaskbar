// dllmain.cpp : Defines the entry point for the DLL application.
//#include "pch.h"

#ifdef HOOKDLL_EXPORTS
#define HOOK_API __declspec(dllexport)
#else
#define HOOK_API __declspec(dllimport)
#endif

#include <windows.h>
#include <tchar.h>

#include <commctrl.h>
#pragma comment (lib, "Comctl32")

#include <Shlobj.h>
#include <shlwapi.h>
#pragma comment (lib, "Shlwapi")

#pragma data_seg ("Shared")

HHOOK g_hHook = NULL;

#pragma data_seg ()

#pragma comment(linker, "/section:Shared,rws")

HINSTANCE g_hInst = NULL;
BOOL bSubclassed = FALSE;
#define WM_PINMESSAGE WM_USER + 10
#define WM_PINNOTIFY WM_USER + 11

LRESULT CALLBACK GetMsgProc(int nCode, WPARAM wParam, LPARAM lParam);
int CompareMenuStrings(LPCWSTR s1, LPCWSTR s2);
HRESULT PinUnpinToTaskbar(LPCWSTR wsPathFile, int* nPinned);

#ifndef IID_PPV_ARG
#define IID_PPV_ARG(IType, ppType) IID_##IType, reinterpret_cast<void**>(static_cast<IType**>(ppType))
#endif

BOOL APIENTRY DllMain( HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
	{
		g_hInst = (HINSTANCE)hModule;	
		//HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	}
	break;
    case DLL_THREAD_ATTACH:
		break;
    case DLL_THREAD_DETACH:
		break;
    case DLL_PROCESS_DETACH:
	{
		//CoUninitialize();
	}
    break;
    }
    return TRUE;
}

// For test
HOOK_API void Test()
{
	MessageBox(NULL, TEXT("Test"), TEXT("Information"), MB_OK | MB_ICONINFORMATION);
}

HOOK_API BOOL SetHook(BOOL bInstall, DWORD dwThreadId)
{
	if (bInstall)
	{	
		g_hHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)GetMsgProc, g_hInst, dwThreadId);
		if (g_hHook)
		{
			Beep(6000, 10);
			/*WCHAR wsMessage[MAX_PATH] = TEXT("");
			wsprintf(wsMessage, TEXT("SetWindowsHookEx: %d - Thread Id: %d"), g_hHook, dwThreadId);
			MessageBox(NULL, wsMessage, TEXT("Information"), MB_OK | MB_ICONINFORMATION);*/
		}
		else
		{
			int nError = GetLastError();
			WCHAR wsMessage[MAX_PATH] = TEXT("");
			wsprintf(wsMessage, TEXT("Error SetWindowsHookEx: %d - Thread Id: %d"), nError, dwThreadId);
			MessageBox(NULL, wsMessage, TEXT("Error"), MB_OK | MB_ICONSTOP);
		}
	}
	else
	{	
		if (g_hHook)
		{
			UnhookWindowsHookEx(g_hHook);
			g_hHook = NULL;
			Beep(1000, 10);
		}
	}
	return TRUE;
}

LRESULT CALLBACK GetMsgProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 0;
	lResult = CallNextHookEx(NULL, nCode, wParam, lParam);
	MSG* lpMsg;
	if ((nCode >= 0) && (PM_REMOVE == wParam))
	{
		lpMsg = (MSG*)lParam;
		if (lpMsg->message == WM_PINMESSAGE)
		{
			//Beep(9000, 10);
			HANDLE hMapFile = OpenFileMapping(FILE_MAP_READ, FALSE, TEXT("PinSharedMemory"));
			if (hMapFile)
			{
				LPVOID pSharedData = MapViewOfFile(hMapFile, FILE_MAP_READ, 0, 0, 0);
				if (pSharedData)
				{					
					WCHAR wsPath[MAX_PATH] = { 0 };
					lstrcpy(wsPath, (WCHAR*)pSharedData);
					//MessageBox(NULL, wsPath, TEXT("Received path from Shared Memory"), MB_OK);
					int nPinned = 0;
					HRESULT hr = PinUnpinToTaskbar(wsPath, &nPinned);
					if (lpMsg->lParam != NULL)
						PostMessage((HWND)lpMsg->lParam, WM_PINNOTIFY, nPinned, hr);
					UnmapViewOfFile(pSharedData);
				}
				CloseHandle(hMapFile);
			}
		}
	}
	return lResult;
}

 //STDAPI SHGetUIObjectFromFullPIDL(LPCITEMIDLIST pidl, HWND hwnd, REFIID riid, void** ppv)
 //{
	// LPCITEMIDLIST pidlChild;
	// IShellFolder* psf;

	// *ppv = NULL;

	// HRESULT hr = SHBindToParent(pidl, IID_PPV_ARG(IShellFolder, &psf), &pidlChild);
	// if (SUCCEEDED(hr))
	// {
	//	 hr = psf->GetUIObjectOf(hwnd, 1, &pidlChild, riid, NULL, ppv);
	//	 psf->Release();
	// }
	// return hr;
 //}

 int CompareMenuStrings(LPCWSTR s1, LPCWSTR s2)
 {
	 while (*s1 != L'\0' && *s2 != L'\0')
	 {
		 // Skip ampersands
		 while (*s1 == L'&') ++s1;
		 while (*s2 == L'&') ++s2;
		 if (towlower(*s1) != towlower(*s2))
			 return (int)(towlower(*s1) - towlower(*s2));
		 ++s1;
		 ++s2;
	 }
	 // Skip trailing ampersands
	 while (*s1 == L'&') ++s1;
	 while (*s2 == L'&') ++s2;
	 return (int)(towlower(*s1) - towlower(*s2));
 }

 BOOL IsAlreadyPinned(HMENU hMenu, LPCWSTR wsUnpinString)
 {
	 BOOL bPinned = FALSE;
	 int nNbItems = GetMenuItemCount(hMenu);
	 for (int i = nNbItems - 1; i >= 0; i--)
	 {
		 MENUITEMINFO mii = {};
		 mii.cbSize = sizeof(mii);
		 mii.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_ID | MIIM_SUBMENU | MIIM_DATA | MIIM_STATE;
		 mii.cch = 128;
		 TCHAR pszName[128] = {};
		 mii.dwTypeData = pszName;
		 if (GetMenuItemInfo(hMenu, i, TRUE, &mii))
		 {			
			 if (CompareMenuStrings(pszName, wsUnpinString) == 0)
			 {
				 bPinned = TRUE;
				 break;
			 }
		 }
	 }
	 return bPinned;
 }

 HRESULT PinUnpinToTaskbar(LPCWSTR wsPathFile, int* nPinned)
 {
	 // https://github.com/fei018/OmegaT-Windows/blob/4a4df43a5dfaf78534d0eeb70258200428591030/Windows/source/8/mui/Windows/System32/be-BY/shell32.dll_6.2.8102.0_x32.rc#L2926

	 *nPinned = 0;
	 WCHAR wszPinLocalizedVerb[128] = { 0 };
	 WCHAR wszUnpinLocalizedVerb[128] = { 0 };
	 HINSTANCE hShell32 = LoadLibrary(TEXT("Shell32.dll"));
	 if (hShell32)
	 {
		 LoadString(hShell32, static_cast<UINT>(5386), wszPinLocalizedVerb, 128);
		 LoadString(hShell32, static_cast<UINT>(5387), wszUnpinLocalizedVerb, 128);
		 FreeLibrary(hShell32);
	 }

	 HRESULT hr = E_FAIL;
	 LPITEMIDLIST pItemIDL = ILCreateFromPath(wsPathFile);
	 if (pItemIDL)
	 {
		 LPCITEMIDLIST aPidl[512];
		 LPITEMIDLIST pidlRel = NULL;
		 UINT nIndex = 0;
		 LPCITEMIDLIST pidlChild;
		 IShellFolder* psf;

		 pidlRel = ILFindLastID(pItemIDL);
		 aPidl[nIndex++] = ILClone(pidlRel);
		 hr = SHBindToParent(pItemIDL, IID_PPV_ARGS(&psf), &pidlChild);
		 if (SUCCEEDED(hr))
		 {
			 IContextMenu* pContextMenu;
			 DEFCONTEXTMENU dcm = { NULL, NULL, NULL, psf, nIndex, (LPCITEMIDLIST*)aPidl, NULL, 0, NULL };
			 hr = SHCreateDefaultContextMenu(&dcm, IID_IContextMenu, (void**)&pContextMenu);
			 if (SUCCEEDED(hr))
			 {
				 HMENU hMenu = CreatePopupMenu();
				 hr = pContextMenu->QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF_NORMAL /*| CMF_EXPLORE | CMF_EXTENDEDVERBS | CMF_ITEMMENU | CMF_SYNCCASCADEMENU | CMF_ASYNCVERBSTATE*/);
				 if (SUCCEEDED(hr))
				 {
					 //if (!IsAlreadyPinned(hMenu, wszUnpinLocalizedVerb))
					 {
						 int nNbItems = GetMenuItemCount(hMenu);
						 for (int i = nNbItems - 1; i >= 0; i--)
						 {
							 MENUITEMINFO mii = {};
							 mii.cbSize = sizeof(mii);
							 mii.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_ID | MIIM_SUBMENU | MIIM_DATA | MIIM_STATE;
							 mii.cch = 128;
							 TCHAR pszName[128] = {};
							 mii.dwTypeData = pszName;
							 if (GetMenuItemInfo(hMenu, i, TRUE, &mii))
							 {
								 // Not in x64 without SHCreateDefaultContextMenu
								 // pszName = "&Épingler à la barre des tâches"
								 // wszLocalizedVerb = "Épingler à la &barre des tâches"
								 //if (lstrcmp(pszName, wszLocalizedVerb) == 0)
								 //if (lstrcmp(pszName, TEXT("P&ropriétés")) == 0)
								 if (CompareMenuStrings(pszName, wszPinLocalizedVerb) == 0 ||
									 CompareMenuStrings(pszName, wszUnpinLocalizedVerb) == 0)
								 {
									 CMINVOKECOMMANDINFOEX invoke = {};
									 invoke.cbSize = sizeof(invoke);
									 invoke.fMask = CMIC_MASK_UNICODE;
									 invoke.hwnd = GetDesktopWindow();
									 invoke.lpVerb = MAKEINTRESOURCEA(mii.wID - 1);
									 invoke.lpVerbW = MAKEINTRESOURCEW(mii.wID - 1);
									 invoke.nShow = SW_SHOWNORMAL;
									 hr = pContextMenu->InvokeCommand((LPCMINVOKECOMMANDINFO)&invoke);
									 if (SUCCEEDED(hr))
									 {
										 if (CompareMenuStrings(pszName, wszPinLocalizedVerb) == 0)
											 *nPinned = 1;
										 else if (CompareMenuStrings(pszName, wszUnpinLocalizedVerb) == 0)
											 *nPinned = 2;
									 }
									/* if (SUCCEEDED(hr))
										 Beep(8000, 10);*/
									 break;
								 }
							 }
						 }
					 }
					/* else
					 {
						 Beep(100, 1000);
					 }*/
				 }
				 DestroyMenu(hMenu);
				 //IUnknown_SetSite(pContextMenu, NULL);
				 pContextMenu->Release();
			 }
			 psf->Release();
		 }
		 ILFree((LPITEMIDLIST)aPidl[0]);
		 ILFree(pItemIDL);
	 }
	 return hr;
 }