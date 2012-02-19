/*
 * This file is part of Open Twebst - web automation framework.
 * Copyright (c) 2012 Adrian Dorache
 * adrian.dorache@codecentrix.com
 *
 * Open Twebst is free software: you can redistribute it
 * and/or modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * Open Twebst is distributed in the hope that it will be
 * useful, but WITHOUT ANY WARRANTY; without even the implied warranty
 * of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with Open Twebst. If not, see <http://www.gnu.org/licenses/>.
 *
 *
 * Twebst can be used under a commercial license if such has been acquired
 * (see http://www.codecentrix.com/). The commercial license does not
 * cover derived or ported versions created by third parties under GPL.
 */

#include "StdAfx.h"
#include "DebugServices.h"
#include "HtmlHelpers.h"
#include "MarshalService.h"
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "..\BrowserPlugin\BrowserPlugin_i.c"
#include "SpinLock.h"
#include "OTWBSTInjector.h"


extern HINSTANCE g_hInstance;
extern UINT      g_nRegisteredMsg;
extern HHOOK     g_hLoaderHook;
extern SpinLock  g_spinLockBho;

static BOOL                       IsIEServerWindow (HWND hTargetWnd);
static BOOL                       OnLoadBHO        (HWND hWnd, BOOL bStopBrowser, CComQIPtr<IWebBrowser2> spBrowser = CComQIPtr<IWebBrowser2>());
static BOOL                       IsBhoLoaded      (HWND hWnd);
static CComQIPtr<IExplorerPlugin> CreateBhoInstance();
static LRESULT CALLBACK           BhoLoaderHookProc(int nCode, WPARAM wParam, LPARAM lParam);



OPENTWBSTINJECTOR_API BOOL InjectBHO(HWND hIEWnd, BOOL bStopBrowser, CComQIPtr<IWebBrowser2> spBrowser)
{
	ATLASSERT(g_nRegisteredMsg != NULL);
	ATLASSERT(g_hInstance      != NULL);

	g_spinLockBho.Lock();

	// Before injecting check it is not already injected.
	DWORD dwIEThread = ::GetWindowThreadProcessId(hIEWnd, NULL);
	if (IsBhoLoaded(hIEWnd))
	{
		g_spinLockBho.Unlock();
		return TRUE;
	}

	// If current thread is the same as dwIEThread then directly load the BHO as we are in the same appartment.
	if (::GetCurrentThreadId() == dwIEThread)
	{
		// This scenario happens when automating an embeded web control.
		traceLog << "Directly load the BHO, no hook required\n";
		BOOL bRet =  OnLoadBHO(hIEWnd, bStopBrowser, spBrowser);

		g_spinLockBho.Unlock();
		return bRet;
	}

	if (!IsIEServerWindow(hIEWnd))
	{
		traceLog << "Invalid hIEWnd in InjectBHO\n";
		g_spinLockBho.Unlock();
		return FALSE;
	}

	HHOOK hHook = ::SetWindowsHookEx(WH_CALLWNDPROC, BhoLoaderHookProc, g_hInstance, dwIEThread);
	if (NULL == hHook)
	{
		g_spinLockBho.Unlock();

		traceLog << "Can not install hook in InjectBHO\n";
		return FALSE;
	}

	g_hLoaderHook = hHook;
	::SendMessage(hIEWnd, g_nRegisteredMsg, NULL, bStopBrowser);

	BOOL bRes = ::UnhookWindowsHookEx(hHook);
	if (FALSE == bRes)
	{
		traceLog << "UnhookWindowsHookEx failed in InjectBHO\n";
	}

	// TODO: see if it is necessary to send another message to actually unload the dll from target process.
	//::SendNotifyMessage(HWND_BROADCAST, WM_STOP_HOOKS, 0, 0);

	g_hLoaderHook = NULL;
	hHook         = NULL;

	g_spinLockBho.Unlock();
	return TRUE;
}


static BOOL IsIEServerWindow(HWND hTargetWnd)
{
	if (!hTargetWnd || !::IsWindow(hTargetWnd))
	{
		return FALSE;
	}
	
	// Retrieve the class name.
	TCHAR	szClassName[MAX_PATH] = { 0 };
	int nClassLength = ::RealGetWindowClass(hTargetWnd, szClassName, MAX_PATH - 1);
	szClassName[nClassLength] = _T('\0');

	// Check if we are on a IE web page
	if(0 == _tcscmp(szClassName, _T("Internet Explorer_Server")))
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


static LRESULT CALLBACK BhoLoaderHookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (!g_hLoaderHook || !g_nRegisteredMsg)
	{
		return 0;
	}

	if (HC_ACTION == nCode)
	{
		CWPSTRUCT* pInfo = (CWPSTRUCT*)lParam;
		if ((pInfo != NULL) && (g_nRegisteredMsg == pInfo->message))
		{
			OnLoadBHO(pInfo->hwnd, (pInfo->lParam != 0));
		}
	}

	return ::CallNextHookEx(g_hLoaderHook, nCode, wParam, lParam);
}


static BOOL OnLoadBHO(HWND hWnd, BOOL bStopBrowser, CComQIPtr<IWebBrowser2> spBrowser)
{
	if (!IsIEServerWindow(hWnd) && !spBrowser)
	{
		traceLog << "Invalid param window in OnLoadBHO\n";
		return FALSE;
	}

	if (!spBrowser)
	{
		spBrowser = HtmlHelpers::GetBrowserFromIEServerWnd(hWnd);
	}

	if (spBrowser != NULL)
	{
		if (bStopBrowser)
		{
			// Get the browswer ready state.
			READYSTATE state;
			HRESULT    hRes  = spBrowser->get_ReadyState(&state);

			if (FAILED(hRes) || ((READYSTATE_COMPLETE != state) && (READYSTATE_UNINITIALIZED != state)))
			{
				// The browser is still loading.
				// Stop the browser so pending download counter is set to zero when initializing the BHO.
				// However it is wise to attach to hosted browser before any other action on it.
				hRes = spBrowser->Stop();
			}
		}

		// Don't call CoCreateInstance; instead directly load the library.
		CComQIPtr<IExplorerPlugin> spBho = CreateBhoInstance();
		if (spBho != NULL)
		{
			CComQIPtr<IObjectWithSite> spSite = spBho;
			if (spSite != NULL)
			{
				HRESULT hRes = spBho->SetForceLoaded(HandleToLong(hWnd));
				if (SUCCEEDED(hRes))
				{
					hRes = spSite->SetSite(spBrowser);					
					if (SUCCEEDED(hRes))
					{
						// Keep the object alive.
						spBho.Detach();
						return TRUE;
					}
				}
			}
		}
	}

	return FALSE;
}


static CComQIPtr<IExplorerPlugin> CreateBhoInstance()
{
	CComQIPtr<IExplorerPlugin> spBho;
	HRESULT hRes = spBho.CoCreateInstance(CLSID_OpenTwebstBHO);
	return spBho;
}


static BOOL IsBhoLoaded(HWND hWnd)
{
	ATLASSERT(IsIEServerWindow(hWnd) || (::GetCurrentThreadId() == ::GetWindowThreadProcessId(hWnd, NULL)));

	String sPluginMarshalName;
	BOOL   bName = NewMarshalService::GetMarshalWndName(hWnd, sPluginMarshalName);

	if (!bName)
	{
		traceLog << "NewMarshalService::GetMarshalWndName failed in IsBhoLoaded\n";
		return FALSE;
	}

	HWND hPluginCommWnd = ::FindWindow(NewMarshalService::HIDDEN_COMMUNICATION_WND_CLASS_NAME, sPluginMarshalName.c_str());
	return (hPluginCommWnd != NULL);
}
