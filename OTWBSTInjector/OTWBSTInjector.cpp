// PutWndInFg.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "DebugServices.h"
#include "OTWBSTInjector.h"
#include "Common.h"
#include "SpinLock.h"
using namespace Common;


#pragma comment(linker, "/section:OTWBSTInjectorSharedSection,rws")
#pragma data_seg("OTWBSTInjectorSharedSection")
	HHOOK g_hHook       = NULL;
	HHOOK g_hLoaderHook = NULL;

	volatile long g_nLock = SL_UNLOCKED;
#pragma data_seg()

const TCHAR REGISTERED_MSG_NAME[] = _T("{B145F25E-E6E9-4ccb-AA66-E373222347FF}");
UINT      g_nRegisteredMsg = 0;
HINSTANCE g_hInstance      = NULL;
SpinLock  g_spinLockBho(&g_nLock);


BOOL OnProcessAttach (HANDLE hModule);
void OnProcessDettach();



///////////////////////////////////////////////////////////////////////////////////
BOOL APIENTRY DllMain(HANDLE hModule,
                      DWORD  ul_reason_for_call,
                      LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		{
			return OnProcessAttach(hModule);
		}

		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		{
			break;
		}

		case DLL_PROCESS_DETACH:
		{
			OnProcessDettach();
			break;
		}
	}

    return TRUE;
}


BOOL OnProcessAttach(HANDLE hModule)
{
	if (NULL == hModule)
	{
		return FALSE;
	}

	g_hInstance = static_cast<HINSTANCE>(hModule);
	g_nRegisteredMsg = ::RegisterWindowMessage(REGISTERED_MSG_NAME);

	if (!g_nRegisteredMsg)
	{
		return FALSE;
	}

	return TRUE;
}


void OnProcessDettach()
{
	g_hInstance      = NULL;
	g_nRegisteredMsg = NULL;
}

LRESULT CALLBACK CallWndProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	CWPSTRUCT* pInfo = (CWPSTRUCT*)lParam;
	if (g_nRegisteredMsg == pInfo->message)
	{
		HWND hWndTarget = (HWND)pInfo->wParam;
		BOOL bRes       = ::SetForegroundWindow(hWndTarget);
		_ASSERTE(bRes != FALSE);
	}

	return ::CallNextHookEx(g_hHook, nCode, wParam, lParam);
}


// Return FALSE if any error occurs.
OPENTWBSTINJECTOR_API BOOL PutWindowInForeground(HWND hTargetWnd)
{
	if (::IsIconic(hTargetWnd))
	{
		::ShowWindow(hTargetWnd, SW_RESTORE);
	}

	HWND hFgWnd = ::GetForegroundWindow();
	if (NULL == hFgWnd)
	{
		// No window in foreground. Set the target wnd in foreground directly.
		BOOL bRes = ::SetForegroundWindow(hTargetWnd);
		return bRes;
	}

	DWORD dwActiveProcess;
	DWORD dwThisProcess = ::GetCurrentProcessId();
	DWORD dwFgThread    = ::GetWindowThreadProcessId(hFgWnd, &dwActiveProcess);

	if (dwActiveProcess == dwThisProcess)
	{
		// The current process is in foreground. Set the target wnd in foreground directly.
		BOOL bRes = ::SetForegroundWindow(hTargetWnd);
		return bRes;
	}

	String sClassName = GetWndClass(hFgWnd);
	if (sClassName == _T("ConsoleWindowClass"))
	{
		// A console window is in foreground. Set the target wnd in foreground directly.
		BOOL bRes = ::SetForegroundWindow(hTargetWnd);
		return bRes;
	}

	g_spinLockBho.Lock();

	g_hHook = ::SetWindowsHookEx(WH_CALLWNDPROC, CallWndProc, g_hInstance, dwFgThread);
	if (NULL == g_hHook)
	{
		traceLog << "Can not install hook in PutWindowInForeground\n";

		g_spinLockBho.Unlock();
		return FALSE;
	}

	::SendMessage(hFgWnd, g_nRegisteredMsg, (WPARAM)hTargetWnd, 0);

	BOOL bRes = ::UnhookWindowsHookEx(g_hHook);
	g_hHook = NULL;

	if (FALSE == bRes)
	{
		traceLog << "UnhookWindowsHookEx failed in PutWindowInForeground\n";

		g_spinLockBho.Unlock();
		return FALSE;
	}

	// TODO: see if it is necessary to send another message to actually unload the dll from target process.
	//::SendNotifyMessage(HWND_BROADCAST, WM_STOP_HOOKS, 0, 0);

	// Wait some time the target window to be in foreground.
	const int PUT_WND_IN_FG_TIMEOUT = 3000;
	DWORD     dwStartWait           = ::GetTickCount();

	while (TRUE)
	{
		HWND hFgWnd = ::GetForegroundWindow();
		if (hFgWnd == hTargetWnd)
		{
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if (dwCurrentTime - dwStartWait > PUT_WND_IN_FG_TIMEOUT)
		{
			break;
		}

		::Sleep(Common::INTERNAL_GLOBAL_PAUSE);
	}

	g_spinLockBho.Unlock();
	return TRUE;
}
