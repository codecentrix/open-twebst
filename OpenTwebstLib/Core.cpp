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

#include "stdafx.h"
#include "HtmlHelpIDs.h"
#include "DebugServices.h"
#include "Exceptions.h"
#include "Registry.h"
#include "CodeErrors.h"
#include "Common.h"
#include "Settings.h"
#include "Core.h"
#include "Frame.h"
#include "Element.h"
#include "Browser.h"
#include "Core.h"
#include "HtmlHelpers.h"
#include "MethodAndPropertyNames.h"
#include "SafeArrayAutoPtr.h"
#include "MarshalService.h"
#include "..\BrowserPlugin\BrowserPlugin_i.c"
#include "SearchCondition.h"


#undef FIRE_CANCEL_REQUEST
#define FIRE_CANCEL_REQUEST() \
{\
	Fire_CancelRequest();\
	if (this->IsCancelPending())\
	{\
		SetComErrorMessage(IDS_ERR_CANCELED, IDH_CORE_CANCELATION);\
		SetLastErrorCode(ERR_CANCELED);\
		return HRES_CANCELED_ERR;\
	}\
}


extern HINSTANCE g_hInstance;

using namespace Registry;
using namespace Common;

struct CallbackWndInfo
{
	CallbackWndInfo()
	{
		m_procId  = 0;
		m_hIeWnd = NULL;
	}

	DWORD m_procId;
	HWND  m_hIeWnd;
};


struct FindIEFrameExcludeSetInfo
{
	FindIEFrameExcludeSetInfo(std::set<HWND>* pIEFramesToExclude) : 
		m_hIEFrameWnd(NULL), m_pIEFramesToExclude(pIEFramesToExclude)
	{
	}

	std::set<HWND>* m_pIEFramesToExclude;
	HWND            m_hIEFrameWnd;
};



STDMETHODIMP CCore::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ICore
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i], riid))
			return S_OK;
	}

	return S_FALSE;
}


HRESULT CCore::Fire_CancelRequest()
{
	if (IsCancelPending())
	{
		// Already canceled.
		return S_OK;
	}

	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();

	for (int iConnection = 0; iConnection < cConnections; iConnection++)
	{
		this->Lock();
		CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
		this->Unlock();

		IDispatch * pConnection = static_cast<IDispatch*>(punkConnection.p);
		if (pConnection)
		{
			VARIANT_BOOL vbCancel = VARIANT_FALSE;
			CComVariant avarParams[1];
			avarParams[0].byref = &vbCancel;
			avarParams[0].vt = VT_BOOL|VT_BYREF;

			DISPPARAMS params = { avarParams, NULL, 1, 0 };
			hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

            if (SUCCEEDED(hr) && (VARIANT_TRUE == vbCancel))
            {
				traceLog << "Cancel requested\n";
				m_bCanceled = TRUE;

                break;
            }
		}
	}

	return hr;
}


BOOL CCore::IsCancelPending()
{
	return m_bCanceled;
}


CCore::CCore()
{
	m_nLastError        = Common::ERR_OK;
	m_bCanceled         = FALSE;
	m_vbAutoClosePopups = VARIANT_FALSE;
	m_bAsyncEvents      = VARIANT_FALSE;

	// Read Core.searchTimeout default value from the registry.
	try
	{
		m_nSearchTimeout = RegGetDWORDValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                    Settings::REG_SETTINGS_SEARCH_TIMEOUT_VALUE);
	}
	catch (const RegistryException&)
	{
		m_nSearchTimeout = Settings::DEFAULT_SEARCH_TIMEOUT;
	}

	// Read Core.loadTimeout default value from the registry.
	try
	{
		m_nLoadTimeout = RegGetDWORDValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                  Settings::REG_SETTINGS_LOAD_TIMEOUT_VALUE);
	}
	catch (const RegistryException&)
	{
		m_nLoadTimeout = Settings::DEFAULT_LOAD_TIMEOUT;
	}

	// Read Core.loadTimeoutIsErr default value from the registry.
	try
	{
		 DWORD dwLoadTimeoutIsErr = RegGetDWORDValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                             Settings::REG_SETTINGS_LOAD_TIMEOUT_IS_ERROR_VALUE);

		 m_vbLoadTimeoutIsErr = (dwLoadTimeoutIsErr ? VARIANT_TRUE : VARIANT_FALSE);
	}
	catch (const RegistryException&)
	{
		m_vbLoadTimeoutIsErr = Settings::DEFAULT_LOAD_TIMEOUT_IS_ERR;
	}

	// Read Core.useIEevents default value from the registry.
	try
	{
		DWORD dwUseIeEvents = RegGetDWORDValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                       Settings::REG_SETTINGS_USE_IE_EVENTS);
		m_vbUseHardwareEvents = (dwUseIeEvents ? VARIANT_TRUE : VARIANT_FALSE);
	}
	catch (const RegistryException&)
	{
		m_vbUseHardwareEvents = Settings::DEFAULT_USE_HARDWARE_EVENTS;
	}

	// Read settings flag from the registry.
	try
	{
		m_nSettingsFlag = RegGetDWORDValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                   Settings::REG_SETTINGS_FLAG_VALUE);
	}
	catch (const RegistryException&)
	{
		m_nSettingsFlag = 0;
	}
}


HRESULT CCore::StartBrowser(IBrowser** ppNewBrowser, DWORD dwHelpID, BOOL bStartHidden, BSTR bstrUrl)
{
	if ((NULL == bstrUrl) || (NULL == ppNewBrowser))
	{
		traceLog << "Invalid URL parameter in CCore::StartBrowser\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	ATLASSERT(ppNewBrowser != NULL);

	CComQIPtr<IExplorerPlugin> spPlugin;
	HWND     hIeWnd = NULL;

	try
	{
		if (!Common::IsWindowsVistaOrLater() && (Common::GetIEVersion() > 7))
		{
			// WinXP and IE8.
			std::set<HWND> allIEFrames;

			FindAllIEFramesWnd(allIEFrames);
			StartInternetExplorerProcess(bStartHidden, bstrUrl);

			hIeWnd = FindIEFrameExcludeSet(allIEFrames);
			if (NULL == hIeWnd)
			{
				traceLog << "FindIEFrameExcludeSet failed in CCore::StartBrowser\n";
				SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);	// Set the error message.
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}
		}
		else
		{
			// Start a new instance of "Internet Explorer" browser and find the top level window.
			DWORD dwIEProcId = StartInternetExplorerProcess(bStartHidden, bstrUrl);

			hIeWnd = FindIeWnd(dwIEProcId);
			if (NULL == hIeWnd)
			{
				traceLog << "FindIeWnd returns NULL in CCore::StartBrowser\n";
				SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);	// Set the error message.
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}
		}

		// Search a browser plugin.
		if (::IsWindow(hIeWnd))
		{
			spPlugin = NewMarshalService::FindBrwsPluginByIeWnd(hIeWnd, Common::INTERNAL_GLOBAL_TIMEOUT);
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		traceLog << "Can NOT start IE browser in CCore::StartBrowser\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (spPlugin == NULL)
	{
		traceLog << "Can NOT start IE browser, NULL returned by MarshalService::FindBrwsPluginByThreadIDInTimeout."
		            " The BHO proxy could not be properly registered\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER_BECAUSE_OF_PROXY, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
	else
	{
		if (bStartHidden)
		{
			// Intentionally do NOT release pWebBrws pointer. Otherwise the hidden IE instance will end.
			IWebBrowser2* pWebBrws = NULL;
			HRESULT       hRes     = spPlugin->GetNativeBrowser(&pWebBrws);
			if (NULL == pWebBrws)
			{
				traceLog << "pWebBrws is NULL. Can NOT start IE browser in CCore::StartBrowser\n";
				SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			if (Common::IsWindowsVistaOrLater() && ::IsWindow(hIeWnd))
			{
				// It must be a top level browser window.
				ATLASSERT(Common::GetWndClass(hIeWnd) == _T("TabWindowClass"));

				// Find top-level IEFrame window.
				HWND hTopWnd = Common::GetTopParentWnd(hIeWnd);

				// Put the web browser window in foreground.
				ATLASSERT(Common::GetWndClass(hTopWnd) == _T("IEFrame"));

				// On Vista hide the main window because there's no way (for now) to start the browser hidden from start.
				::ShowWindow(hTopWnd, SW_HIDE);
			}
		}

		// Create a new Browser object.
		IBrowser* pBrws = NULL;
		HRESULT   hRes = CComCoClass<CBrowser>::CreateInstance(&pBrws);
		if (pBrws != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));

			CBrowser* pBrowser = static_cast<CBrowser*>(pBrws);	// Down cast !!!
			ATLASSERT(pBrowser != NULL);
			pBrowser->SetPlugin(spPlugin);
			pBrowser->SetCore(this);

			// Set the object reference into the output pointer.
			*ppNewBrowser = pBrws;
			return HRES_OK;
		}

		traceLog << "Can not create Browser instance in CCore::StartBrowser\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_START_BROWSER, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
}


STDMETHODIMP CCore::StartBrowser(BSTR bstrUrl, IBrowser** ppNewBrowser)
{
	FIRE_CANCEL_REQUEST();
	return StartBrowser(ppNewBrowser, IDH_CORE_START_BROWSER, FALSE, bstrUrl);
}


STDMETHODIMP CCore::FindBrowser(BSTR bstrCond, IBrowser** ppBrowserFound)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pArguments;
	pArguments.AddMultiCondition(bstrCond);

	int nSafeArraySize;
	if ((NULL == pArguments) || !Common::SafeArraySize(pArguments, &nSafeArraySize))
	{
		traceLog << "Invalid SAFEARRAY parameter in CCore::FindBrowser\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, FIND_BROWSER_METHOD, IDH_CORE_FIND_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	SAFEARRAY*        pVarArgs;
	SafeArrrayAutoPtr newSafeArray;

	BOOL   bUseEqOp         = TRUE;
	String sAppName         = _T("IEXPLORE.EXE");
	BOOL   bAddAnyPidFilter = FALSE;

	if (nSafeArraySize != 0)
	{
		pVarArgs = pArguments;

		std::list<DescriptorToken> tokenList;
		BOOL bValidArgs = Common::GetDescriptorTokensList(pVarArgs, &tokenList);
		if (!bValidArgs || !IsValidDescriptorList(tokenList))
		{
			traceLog << "Invalid list of descriptor tokens in CCore::FindBrowser\n";
			SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, FIND_BROWSER_METHOD, IDH_CORE_FIND_BROWSER);
			SetLastErrorCode(ERR_INVALID_ARG);
			return HRES_INVALID_ARG;
		}

		if (GetAppFilterData(tokenList, sAppName, bUseEqOp))
		{
			if (tokenList.size() == 1)
			{
				bAddAnyPidFilter = TRUE;
			}
		}
	}
	else
	{
		bAddAnyPidFilter = TRUE;
	}

	if (bAddAnyPidFilter)
	{
		// Create a new safe array.
		BOOL           bFail       = FALSE;
		SAFEARRAYBOUND rgsabound   = { 1, 0 };
		SAFEARRAY*     pNewVarArgs = ::SafeArrayCreate(VT_VARIANT, 1, &rgsabound);

		if (pNewVarArgs != NULL)
		{
			CComVariant varNew(L"pid=*");
			long        nIndex = 0;
			HRESULT     hRes   = ::SafeArrayPutElement(pNewVarArgs, &nIndex, &varNew);
			if (SUCCEEDED(hRes))
			{
				pVarArgs = pNewVarArgs;
				newSafeArray.Attach(pNewVarArgs);	// To auto destroy on function exit.
			}
			else
			{
				bFail = TRUE;
			}
		}
		else
		{
			bFail = TRUE;
		}

		if (bFail)
		{
			traceLog << "bFail is true in CCore::FindBrowser\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, FIND_BROWSER_METHOD, IDH_CORE_FIND_BROWSER);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	CComQIPtr<IBrowser> spBrowser;
	try
	{
		DWORD dwStartTime = ::GetTickCount();
		while (TRUE)
		{
			FIRE_CANCEL_REQUEST();

			// Try to find the browser.
			spBrowser = FindBrowser(pVarArgs, sAppName, bUseEqOp);
			if (spBrowser != NULL)
			{
				// The browser was found.
				break;
			}

			DWORD dwCurrentTime = ::GetTickCount();
			if ((dwCurrentTime - dwStartTime) > (m_nSearchTimeout * TIME_SCALE))
			{
				// Timeout has expired. Quit the loop.
				break;
			}

			// Sleep for a while.
			ProcessMessagesOrSleep();
		}
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, FIND_BROWSER_METHOD, IDH_CORE_FIND_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		traceLog << "CCore::FindBrowser failed\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, FIND_BROWSER_METHOD, IDH_CORE_FIND_BROWSER);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	*ppBrowserFound = spBrowser.Detach();
	return HRES_OK;
}


// Returns FALSE is there are some browser still loading.
// Throw an Exception::Services::Exception in case of an error. 
BOOL CCore::FindBrowserList
	(
		SAFEARRAY*                                       pVarArgs,
		std::list<CAdapt<CComQIPtr<IExplorerPlugin> > >* pBrowserList,
		const String&                                    sAppName,
		BOOL                                             bUseEqOp
	)
{
	ATLASSERT(pVarArgs     != NULL);
	ATLASSERT(pBrowserList != NULL);
	ATLASSERT(pBrowserList->empty());

	std::list<HWND> ieWndList;
	NewMarshalService::GetIEServerList(ieWndList, sAppName, bUseEqOp);

	BOOL bResult = TRUE;
	for (std::list<HWND>::iterator it = ieWndList.begin(); it != ieWndList.end(); ++it)
	{
		CComQIPtr<IExplorerPlugin> spCurrentExplorerPlugin = NewMarshalService::FindBrwsPluginByIeWnd(*it);
		if (spCurrentExplorerPlugin != NULL)
		{
			LONG    nSearchFlags = 0;
			HRESULT hRes         = spCurrentExplorerPlugin->CheckBrowserDescriptor(nSearchFlags, pVarArgs);
			if (S_OK == hRes)
			{
				// A browser was found.
				VARIANT_BOOL vbIsLoading;
				hRes = spCurrentExplorerPlugin->IsLoading(&vbIsLoading);
				if (FAILED(hRes))
				{
					traceLog << "spCurrentExplorerPlugin->IsLoading failed in CBrowser::FindBrowserList continue with next browser\n";
					continue;
				}

				if (VARIANT_TRUE == vbIsLoading)
				{
					bResult = FALSE;
				}

				// Add the browser to the list.
				pBrowserList->push_back(spCurrentExplorerPlugin);
			}
		}
	}

	return bResult;
}


// Returns a smart pointer to a browser object or NULL if the browser was not found.
// Could throw a ExceptionServices::Exception in case of major failure.
CComQIPtr<IBrowser> CCore::FindBrowser
	(
		SAFEARRAY*    pVarArgs,
		const String& sAppName,
		BOOL          bUseEqOp
	)
{
	ATLASSERT(pVarArgs != NULL);

	std::list<HWND> ieWndList;
	NewMarshalService::GetIEServerList(ieWndList, sAppName, bUseEqOp);

	CComQIPtr<IExplorerPlugin> spExplorerPlugin;
	for (std::list<HWND>::iterator it = ieWndList.begin(); it != ieWndList.end(); ++it)
	{
		CComQIPtr<IExplorerPlugin> spCurrentExplorerPlugin = NewMarshalService::FindBrwsPluginByIeWnd(*it);
		if (spCurrentExplorerPlugin != NULL)
		{
			LONG    nSearchFlags = 0;
			HRESULT hRes         = spCurrentExplorerPlugin->CheckBrowserDescriptor(nSearchFlags, pVarArgs);

			if (S_OK == hRes)
			{
				// The browser was found. Exit the loop.
				spExplorerPlugin = spCurrentExplorerPlugin;
				break;
			}
			else if (E_INVALIDARG == hRes)
			{
				throw CreateInvalidParamException(_T("Invalid search condition in FindBrowser\n"));
			}
		}
	}

	if (spExplorerPlugin != NULL)
	{
		CComQIPtr<IBrowser> spBrws;
		HRESULT hRes = CComCoClass<CBrowser>::CreateInstance(&spBrws);
		if (spBrws != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));

			CBrowser* pBrowser = static_cast<CBrowser*>((IBrowser*)spBrws);	// Down cast !!!
			ATLASSERT(pBrowser != NULL);
			pBrowser->SetPlugin(spExplorerPlugin);
			pBrowser->SetCore(this);

			return spBrws;
		}
		else
		{
			throw CreateException(_T("Can NOT create a Browser object in CCore::FindBrowser\n"));
		}
	}
	else
	{
		// The browser was not found. Returns NULL.
		return CComQIPtr<IBrowser>();
	}
}


HWND CCore::FindIeWndOnVistaIE8(DWORD dwIEProcID)
{
	ATLASSERT(dwIEProcID != 0);
	ATLASSERT((Common::GetIEVersion() > 7) && Common::IsWindowsVistaOrLater());

	CallbackWndInfo wndInfo;
	wndInfo.m_procId = dwIEProcID;

	DWORD dwStartTime = ::GetTickCount();
	while(TRUE)
	{
		// Enum top level windows seraching a "IEFrame" window that contains a tab that belongs to dwIEProcID process.
		::EnumWindows(FindWndIE8OnVistaCallback, (LPARAM)(&wndInfo));
		if (wndInfo.m_hIeWnd != NULL)
		{
			// We have found the window.
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		ProcessMessagesOrSleep();
	}

	return wndInfo.m_hIeWnd;
}


HWND CCore::FindIeWnd(DWORD dwIEProcID)
{
	// For IE8 on Vista or later dwIEProcID is the id of the tab process.
	// It seems that each tab has its own process.
	if ((Common::GetIEVersion() > 7) && Common::IsWindowsVistaOrLater())
	{
		// For IE8 on Vista or later.
		return FindIeWndOnVistaIE8(dwIEProcID);
	}

	traceLog << "CCore::FindIeWnd begins\n";
	ATLASSERT(dwIEProcID != 0);

	CallbackWndInfo wndInfo;
	wndInfo.m_procId = dwIEProcID;

	DWORD dwStartTime = ::GetTickCount();
	while(TRUE)
	{
		// Enum top level windows seraching a "IEFrame" window that belongs to dwTid thread.
		::EnumWindows(FindWndCallback, (LPARAM)(&wndInfo));
		if (wndInfo.m_hIeWnd != NULL)
		{
			// We have found the window.
			traceLog << "CCore::FindIeWnd windows found: " << wndInfo.m_hIeWnd << "\n";
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
		{
			// Timeout has expired. Quit the loop.
			traceLog << "CCore::FindIeWnd timeout\n";
			break;
		}

		ProcessMessagesOrSleep();
	}

	if (Common::GetIEVersion() > 6)
	{
		// IE7 or IE8 on XP.
		if (wndInfo.m_hIeWnd != NULL)
		{
			dwStartTime = ::GetTickCount();
			while (TRUE)
			{
				// Find a tab.
				HWND hTabWnd = FindTabWnd(wndInfo.m_hIeWnd);
				if (hTabWnd != NULL)
				{
					// Tab window found.
					return hTabWnd;
				}

				DWORD dwCurrentTime = ::GetTickCount();
				if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
				{
					// Timeout has expired. Quit the loop.
					break;
				}

				ProcessMessagesOrSleep();
			}
		}

		return NULL;
	}
	else
	{
		traceLog << "CCore::FindIeWnd ret=" << wndInfo.m_hIeWnd << "\n";
		return wndInfo.m_hIeWnd;
	}
}


BOOL CALLBACK CCore::FindTabWndCallback(HWND hWnd, LPARAM lParam)
{
	ATLASSERT(::IsWindow(hWnd));
	ATLASSERT(lParam != NULL);

	if (Common::GetWndClass(hWnd) == _T("TabWindowClass"))
	{
		HWND* pWnd = (HWND*)lParam;
		*pWnd = hWnd;
		return FALSE;
	}

	return TRUE;
}


HWND CCore::FindTabWnd(HWND hIEFrameWnd)
{
	ATLASSERT(Common::GetWndClass(hIEFrameWnd) == _T("IEFrame"));
	HWND hTabWnd = NULL;
	::EnumChildWindows(hIEFrameWnd, FindTabWndCallback, (LPARAM)(&hTabWnd));

	return hTabWnd;
}


// Starts a new instance of "Internet Explorer" browser and return the process id
// It could throw a RegistryException or Exception in case of error.
DWORD CCore::StartInternetExplorerProcess(BOOL bStartHidden, BSTR bstrUrl)
{
	ATLASSERT(bstrUrl != NULL);

	if (Common::IsWindowsVistaOrLater())
	{
		return StartInternetExplorerProcessOnVista(bstrUrl, bStartHidden);
	}
	else
	{
		traceLog << "CCore::StartInternetExplorer begins\n";

		// Get the path to "Internet Explorer" borwser from registry.
		String sIePath = RegGetStringValue(HKEY_LOCAL_MACHINE,
										   _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\IEXPLORE.EXE"),
										   _T(""));

		// Create the "Internet Explorer" process.
		PROCESS_INFORMATION	pi;
		STARTUPINFO			si;

		::ZeroMemory(&si, sizeof(STARTUPINFO));
		si.cb          = sizeof(STARTUPINFO);
		si.dwFlags     = STARTF_USESHOWWINDOW;
		si.wShowWindow = (bStartHidden ? SW_HIDE : SW_SHOWNORMAL);

		// Compute command line based on exe name and start url.
		USES_CONVERSION;
		String sCommandLine = _T("iexplore.exe ");
		sCommandLine += W2T(bstrUrl);

		BOOL bResult = ::CreateProcess(sIePath.c_str(), (LPTSTR)sCommandLine.c_str(),
		                               NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
		if (!bResult)
		{
			throw CreateException(_T("Can NOT create IE process in CCore::StartInternetExplorer"));
		}

		// Wait for IE process to finish the initialization and starts to wait for user input.
		DWORD dwWaitTimeout = 60000;	// 1 minute should be enough.
		DWORD dwResult      = ::WaitForInputIdle(pi.hProcess, dwWaitTimeout);
		if (dwResult != 0)
		{
			::CloseHandle(pi.hProcess);
			::CloseHandle(pi.hThread);
			throw CreateException(_T("WaitForInputIdle failed in CCore::StartInternetExplorer"));
		}

		DWORD dwProcId = pi.dwProcessId;
		::CloseHandle(pi.hProcess);
		::CloseHandle(pi.hThread);

		traceLog << "CCore::StartInternetExplorer ret=" << dwProcId << "\n";
		return dwProcId;
	}
}


DWORD CCore::StartInternetExplorerProcessOnVista(LPCWSTR pszUrl, BOOL bStartHidden)
{
	UNREFERENCED_PARAMETER(bStartHidden);
	ATLASSERT(pszUrl != NULL);

	traceLog << "CCore::StartInternetExplorerProcessOnVista begins\n";
	ATLASSERT(Common::IsWindowsVistaOrLater());

	HMODULE hModule = ::LoadLibrary(_T("ieframe.dll"));
	if (NULL == hModule)
	{
		throw CreateException(_T("Failed to LoadLibrary(ieframe.dll) in StartInternetExplorerProcessOnVista\n"));
	}

	typedef struct _IELAUNCHURLINFO
	{
		DWORD cbSize;
		DWORD dwCreationFlags;
	} IELAUNCHURLINFO, *LPIELAUNCHURLINFO;

	typedef HRESULT (WINAPI *IELaunchURLProcedure)(LPCWSTR, LPPROCESS_INFORMATION, LPIELAUNCHURLINFO);
	IELaunchURLProcedure ieLauchProc = (IELaunchURLProcedure)::GetProcAddress(hModule, "IELaunchURL");

	DWORD dwProcId = 0;
	if (ieLauchProc != NULL)
	{
		PROCESS_INFORMATION procInfo     = { 0 };
		//IELAUNCHURLINFO     ieLaunchInfo = { 0 };
		//ieLaunchInfo.cbSize = sizeof(IELAUNCHURLINFO);

		HRESULT hRes = ieLauchProc(pszUrl, &procInfo, NULL);
		if (hRes != S_OK)
		{
			::FreeLibrary(hModule);
			hModule = NULL;

			throw CreateException(_T("IELaunchURL failed in StartInternetExplorerProcessOnVista\n"));
		}

		// Wait for IE process to finish the initialization and starts to wait for user input.
		DWORD dwWaitTimeout = 60000;	// 1 minute should be enough.
		DWORD dwResult      = ::WaitForInputIdle(procInfo.hProcess, dwWaitTimeout);
		if (dwResult != 0)
		{
			::CloseHandle(procInfo.hProcess);
			::CloseHandle(procInfo.hThread);
			::FreeLibrary(hModule);
			hModule = NULL;

			throw CreateException(_T("WaitForInputIdle failed in CCore::StartInternetExplorerProcessOnVista"));
		}

		dwProcId = procInfo.dwProcessId;
		::CloseHandle(procInfo.hProcess);
		::CloseHandle(procInfo.hThread);
	}
	else
	{
		::FreeLibrary(hModule);
		hModule = NULL;

		throw CreateException(_T("Failed to GetProcAddress(IELaunchURL) in StartInternetExplorerProcessOnVista\n"));
	}

	::FreeLibrary(hModule);
	hModule = NULL;

	traceLog << "CCore::StartInternetExplorerProcessOnVista end\n";
	return dwProcId;
}


BOOL CALLBACK CCore::FindTabWndIE8OnVistaCallback(HWND hWnd, LPARAM lParam)
{
	ATLASSERT(::IsWindow(hWnd));
	ATLASSERT(lParam != NULL);

	if (Common::GetWndClass(hWnd) == _T("TabWindowClass"))
	{
		CallbackWndInfo* pWndInfo = reinterpret_cast<CallbackWndInfo*>(lParam);
		ATLASSERT(pWndInfo != NULL);

		DWORD dwPid = 0;
		::GetWindowThreadProcessId(hWnd, &dwPid);

		if (dwPid == pWndInfo->m_procId)
		{
			// Stop search we have found the tab window.
			pWndInfo->m_hIeWnd = hWnd;
			return FALSE;
		}
	}

	return TRUE;
}


void CCore::FindTabWndIE8OnVista(HWND hIEFrameWnd, LPARAM lParam)
{
	ATLASSERT(lParam != NULL);
	ATLASSERT(Common::GetWndClass(hIEFrameWnd) == _T("IEFrame"));

	::EnumChildWindows(hIEFrameWnd, FindTabWndIE8OnVistaCallback, lParam);
}


// Enum all top level IE windows of class "IEFrame".
// For each "IEFrame" find a tab that belongs to pWndInfo->m_procId process.
BOOL CALLBACK CCore::FindWndIE8OnVistaCallback(HWND hWnd, LPARAM lParam)
{
	if (GetWndClass(hWnd) == _T("IEFrame"))
	{
		FindTabWndIE8OnVista(hWnd, lParam);

		CallbackWndInfo* pWndInfo = reinterpret_cast<CallbackWndInfo*>(lParam);
		ATLASSERT(pWndInfo != NULL);

		if (pWndInfo->m_hIeWnd != NULL)
		{
			// We have found the window. Stop searching.
			return FALSE;
		}
	}

	// Continue searching.
	return TRUE;
}


// Find a top level IE window that is of class "IEFrame".
// lParam is a pointer to a CallbackWndInfo structure that contains the thread id
// for wich we search the window.
BOOL CALLBACK CCore::FindWndCallback(HWND hWnd, LPARAM lParam)
{
	CallbackWndInfo* pWndInfo = reinterpret_cast<CallbackWndInfo*>(lParam);
	ATLASSERT(pWndInfo != NULL);

	DWORD dwPid = 0;
	::GetWindowThreadProcessId(hWnd, &dwPid);

	// Check only the windows that belong to the given thread.
	if (pWndInfo->m_procId == dwPid)
	{
		if (GetWndClass(hWnd) != _T("IEFrame"))
		{
			// It is not a top level "Internet Explorer" window.
			return TRUE;
		}

		// We have found the window. Stop searching.
		pWndInfo->m_hIeWnd = hWnd;
		return FALSE;
	}

	return TRUE;
}


STDMETHODIMP CCore::get_searchTimeout(LONG* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_searchTimeout\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, SEARCH_TIMEOUT_PROPERTY, IDH_CORE_SEARCH_TIMEOUT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = m_nSearchTimeout;
	return HRES_OK;
}


STDMETHODIMP CCore::put_searchTimeout(LONG newVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (m_nSettingsFlag & Settings::USE_DEFAULT_SEARCH_TIMEOUT_FLAG)
	{
		traceLog << "Use searchTimeout default value in CCore::put_searchTimeout\n";
	}
	else
	{
		m_nSearchTimeout = newVal;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::get_loadTimeout(LONG* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_loadTimeout\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, LOAD_TIMEOUT_PROPERTY, IDH_CORE_LOAD_TIMEOUT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = m_nLoadTimeout;
	return HRES_OK;
}


STDMETHODIMP CCore::put_loadTimeout(LONG newVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (m_nSettingsFlag & Settings::USE_DEFAULT_LOAD_TIMEOUT_FLAG)
	{
		traceLog << "Use loadTimeout default value in CCore::put_loadTimeout\n";
	}
	else
	{
		m_nLoadTimeout = newVal;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::get_lastError(LONG* pVal)
{
	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_lastError\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, LAST_ERR_PROPERTY, IDH_CORE_LAST_ERROR);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = m_nLastError;
	return HRES_OK;
}


void CCore::SetLastErrorCode(LONG nErr)
{
	m_nLastError = nErr;
}


STDMETHODIMP CCore::get_loadTimeoutIsError(VARIANT_BOOL* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_loadTimeoutIsError\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, LOAD_TIMEOUT_IS_ERR_PROPERTY, IDH_CORE_LOAD_TIMEOUT_IS_ERR);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = m_vbLoadTimeoutIsErr;
	return HRES_OK;
}


STDMETHODIMP CCore::put_loadTimeoutIsError(VARIANT_BOOL newVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((newVal != VARIANT_FALSE) && (newVal != VARIANT_TRUE))
	{
		traceLog << "Invalid newVal parameter in CCore::put_loadTimeoutIsError\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, LOAD_TIMEOUT_IS_ERR_PROPERTY, IDH_CORE_LOAD_TIMEOUT_IS_ERR);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_nSettingsFlag & Settings::USE_DEFAULT_LOAD_TIMEOUT_IS_ERROR_FLAG)
	{
		traceLog << "Use loadTimeoutIsError default value in CCore::put_loadTimeoutIsError\n";
	}
	else
	{
		m_vbLoadTimeoutIsErr = newVal;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::get_OK_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_OK_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_OK;
	return HRES_OK;
}


STDMETHODIMP CCore::get_FAIL_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_FAIL_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_FAIL;
	return HRES_OK;
}


STDMETHODIMP CCore::get_INVALID_ARG_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_INVALID_ARG_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_INVALID_ARG;
	return HRES_OK;
}


STDMETHODIMP CCore::get_LOAD_TIMEOUT_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_LOAD_TIMEOUT_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_LOAD_TIMEOUT;
	return HRES_OK;
}


STDMETHODIMP CCore::get_INDEX_OUT_OF_BOUND_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_INDEX_OUT_OF_BOUND_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_INDEX_OUT_OF_BOUND;
	return HRES_OK;
}


STDMETHODIMP CCore::get_BROWSER_CONNECTION_LOST_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_BROWSER_CONNECTION_LOST_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_BRWS_CONNECTION_LOST;
	return HRES_OK;
}


STDMETHODIMP CCore::get_INVALID_OPERATION_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_INVALID_OPERATION_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_OPERATION_NOT_APPLICABLE;
	return HRES_OK;
}


STDMETHODIMP CCore::get_NOT_FOUND_ERROR(LONG* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_NOT_FOUND_ERROR\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ERROR_CONSTANT, IDH_CORE_CONSTANTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = ERR_NOT_FOUND;
	return HRES_OK;
}


HRESULT CCore::SetComErrorMessage(UINT nId, DWORD dwHelpID)
{
	String sHelpFile;
	try
	{
		// Get the help file full path name.
		sHelpFile = Registry::RegGetStringValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                        Settings::REG_SETTINGS_INSTALLATION_DIRECTORY);
		if (_T('\\') != sHelpFile[sHelpFile.size() - 1])
		{
			sHelpFile += _T('\\');
		}

		USES_CONVERSION;
		sHelpFile += Settings::HELP_FILE_RELATIVE_PATH;
		return Error(nId, dwHelpID, T2W((LPTSTR)sHelpFile.c_str()));
	}
	catch (const RegistryException&)
	{
		// Can not get the help file full path name.
		return Error(nId);
	}
}


HRESULT CCore::SetComErrorMessage(UINT nIdPattern, LPCTSTR szFunctionName, DWORD dwHelpID)
{
	// Compute the error message.
	ATLASSERT(g_hInstance != NULL);
	TCHAR szPatternBuffer[128];
	int nRes = ::LoadString(g_hInstance, nIdPattern, szPatternBuffer, (sizeof(szPatternBuffer) / sizeof(szPatternBuffer[0])));
	ATLASSERT(nRes);

#pragma warning(disable:4996)
	TCHAR szErrorMsg[512];
	_stprintf(szErrorMsg, szPatternBuffer, szFunctionName);
#pragma warning(default:4996)

	String sHelpFile;
	try
	{
		// Get the help file full path name.
		sHelpFile = Registry::RegGetStringValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                        Settings::REG_SETTINGS_INSTALLATION_DIRECTORY);
		if (_T('\\') != sHelpFile[sHelpFile.size() - 1])
		{
			sHelpFile += _T('\\');
		}

		sHelpFile += Settings::HELP_FILE_RELATIVE_PATH;
		return Error(szErrorMsg, dwHelpID, sHelpFile.c_str());
	}
	catch (const RegistryException&)
	{
		// Can not get the help file full path name.
		return Error(szErrorMsg);
	}
}


CComQIPtr<IExplorerPlugin> CCore::FindBrwsPluginFromHtmlElement(IHTMLElement* pHtmlElement)
{
	ATLASSERT(pHtmlElement != NULL);

	// Get the document of the html element.
	CComQIPtr<IDispatch> spDispatch;
	HRESULT hRes = pHtmlElement->get_document(&spDispatch);

	CComQIPtr<IHTMLDocument2> spDocument = spDispatch;
	if ((hRes != S_OK) || (spDocument == NULL))
	{
		traceLog << "Can not get the document from element in CCore::FindBrwsPluginFromHtmlElement\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	// Get the html window from document.
	CComQIPtr<IHTMLWindow2> spWindow;
	hRes = spDocument->get_parentWindow(&spWindow);
	if ((hRes != S_OK) || (spWindow == NULL))
	{
		traceLog << "Can not get the window from document in CCore::FindBrwsPluginFromHtmlElement\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	return FindBrwsPluginFromHtmlWindow(spWindow);
}


CComQIPtr<IExplorerPlugin> CCore::FindBrwsPluginFromHtmlWindow(IHTMLWindow2* pHtmlWindow)
{
	ATLASSERT(pHtmlWindow != NULL);

	// Get the top-most IHTMLWindow2 object.
	CComQIPtr<IHTMLWindow2> spTopWindow;
	HRESULT hRes = pHtmlWindow->get_top(&spTopWindow);
	if ((hRes != S_OK) || (spTopWindow == NULL))
	{
		traceLog << "Can not get the top window in CCore::FindBrwsPluginFromHtmlWindow\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	// Get the IWebBrowser2 object.
	CComQIPtr<IWebBrowser2> spBrowser = HtmlHelpers::HtmlWindowToHtmlWebBrowser(spTopWindow);
	if (spBrowser == NULL)
	{
		traceLog << "Can not get the IWebBrowser2 in CCore::FindBrwsPluginFromHtmlWindow\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	return FindBrwsPluginFromWebBrowser(spBrowser);
}


CComQIPtr<IExplorerPlugin> CCore::FindBrwsPluginFromWebBrowser(IWebBrowser2* pWebBrowser)
{
	ATLASSERT(pWebBrowser != NULL);

	// Get the "TabWindowClass" for IE7, "IEFrame" for IE6.
	HWND hIEWnd = HtmlHelpers::GetIEWndFromBrowser(pWebBrowser);
	if (!::IsWindow(hIEWnd))
	{
		return CComQIPtr<IExplorerPlugin>();
	}

	// Find the BHO based on thread ID.
	CComQIPtr<IExplorerPlugin> spPlugin = NewMarshalService::FindBrwsPluginByIeWnd(hIEWnd, 0, TRUE, pWebBrowser);
	if (spPlugin == NULL)
	{
		traceLog << "FindBrwsPluginByIeWnd returned NULL in CCore::FindBrwsPluginFromWebBrowser\n";
	}

	return spPlugin;
}


STDMETHODIMP CCore::AttachToNativeFrame(IHTMLWindow2* pHtmlWindow, IFrame** ppFrame)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	BOOL bInvalidArg = FALSE;
	if (NULL == pHtmlWindow)
	{
		traceLog << "pHtmlWindow is NULL in CCore::AttachToNativeFrame\n";
		bInvalidArg = TRUE;
	}
	else if (NULL == ppFrame)
	{
		traceLog << "ppFrame is NULL in CCore::AttachToNativeFrame\n";
		bInvalidArg = TRUE;
	}

	if (bInvalidArg)
	{
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_FRAME_METHOD, IDH_CORE_CREATE_FRAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	try
	{
		CComQIPtr<IExplorerPlugin> spPlugin = FindBrwsPluginFromHtmlWindow(pHtmlWindow);
		if (spPlugin == NULL)
		{
			throw CreateException(_T("Can NOT find plug-in in ROT in CCore::AttachToNativeFrame\n"));
		}

		// Create a new frame object.
		IFrame* pNewFrame = NULL;
		HRESULT hRes      = CComCoClass<CFrame>::CreateInstance(&pNewFrame);
		if (pNewFrame != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));
			CFrame* pFrameObject = static_cast<CFrame*>(pNewFrame); // Down cast !!!
			ATLASSERT(pFrameObject != NULL);
			pFrameObject->SetPlugin(spPlugin);
			pFrameObject->SetCore(this);
			pFrameObject->SetHtmlWindow(pHtmlWindow);

			// Set the object reference into the output pointer.
			*ppFrame = pNewFrame;
		}
		else
		{
			traceLog << "Can not create a Frame object in CCore::AttachToNativeFrame\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_CORE_CREATE_FRAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ATTACH_TO_FRAME_METHOD, IDH_CORE_CREATE_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::AttachToNativeElement(IHTMLElement* pHtmlElement, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	BOOL bInvalidArg = FALSE;
	if (NULL == pHtmlElement)
	{
		traceLog << "pHtmlElement is NULL in CCore::AttachToNativeElement\n";
		bInvalidArg = TRUE;
	}
	else if (NULL == ppElement)
	{
		traceLog << "ppElement is NULL in CCore::AttachToNativeElement\n";
		bInvalidArg = TRUE;
	}

	if (bInvalidArg)
	{
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_ELEMENT_METHOD, IDH_CORE_CREATE_ELEMENT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	try
	{
		CComQIPtr<IExplorerPlugin> spPlugin = FindBrwsPluginFromHtmlElement(pHtmlElement);
		if (spPlugin == NULL)
		{
			throw CreateException(_T("Can NOT find plug-in in ROT in CCore::AttachToNativeElement\n"));
		}

		// Create a new element object.
		IElement* pNewElement = NULL;
		HRESULT hRes        = CComCoClass<CElement>::CreateInstance(&pNewElement);
		if (pNewElement != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));
			CElement* pElementObject = static_cast<CElement*>(pNewElement); // Down cast !!!
			ATLASSERT(pElementObject != NULL);
			pElementObject->SetPlugin(spPlugin);
			pElementObject->SetCore(this);
			pElementObject->SetHtmlElement(pHtmlElement);

			// Set the object reference into the output pointer.
			*ppElement = pNewElement;
		}
		else
		{
			traceLog << "Can not create a Element object in CCore::AttachToNativeElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_CORE_CREATE_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ATTACH_TO_ELEMENT_METHOD, IDH_CORE_CREATE_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::get_useHardwareInputEvents(VARIANT_BOOL* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CCore::get_useHardwareInputEvents\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, USE_IE_INPUT_EVENTS_PROPERTY, IDH_CORE_USE_IE_EVENTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	*pVal = m_vbUseHardwareEvents;
	return HRES_OK;
}


STDMETHODIMP CCore::put_useHardwareInputEvents(VARIANT_BOOL newVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((newVal != VARIANT_FALSE) && (newVal != VARIANT_TRUE))
	{
		traceLog << "Invalid newVal parameter in CCore::put_useHardwareInputEvents\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, USE_IE_INPUT_EVENTS_PROPERTY, IDH_CORE_USE_IE_EVENTS);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_nSettingsFlag & Settings::USE_DEFAULT_IE_EVENTS_FLAG)
	{
		traceLog << "Use useIEevents default value in CCore::put_useHardwareInputEvents\n";
	}
	else
	{
		m_vbUseHardwareEvents = newVal;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::AttachToNativeBrowser(IWebBrowser2* pWebBrowser, IBrowser** ppBrowser)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	BOOL bInvalidArg = FALSE;
	if (NULL == pWebBrowser)
	{
		traceLog << "pWebBrowser is NULL in CCore::AttachNativeBrowser\n";
		bInvalidArg = TRUE;
	}
	else if (NULL == ppBrowser)
	{
		traceLog << "ppBrowser is NULL in CCore::AttachNativeBrowser\n";
		bInvalidArg = TRUE;
	}

	if (bInvalidArg)
	{
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_BROWSER, IDH_CORE_CREATE_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	try
	{
		CComQIPtr<IExplorerPlugin> spPlugin = FindBrwsPluginFromWebBrowser(pWebBrowser);
		if (spPlugin == NULL)
		{
			throw CreateException(_T("Can NOT find plug-in in ROT in CCore::AttachNativeBrowser\n"));
		}

		// Create a new browser object.
		IBrowser* pNewBrowser = NULL;
		HRESULT   hRes      = CComCoClass<CBrowser>::CreateInstance(&pNewBrowser);
		if (pNewBrowser != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));
			CBrowser* pBrowserObject = static_cast<CBrowser*>(pNewBrowser); // Down cast !!!
			ATLASSERT(pBrowserObject != NULL);
			pBrowserObject->SetPlugin(spPlugin);
			pBrowserObject->SetCore(this);

			// Set the object reference into the output pointer.
			*ppBrowser = pNewBrowser;
		}
		else
		{
			traceLog << "Can not create a Browser object in CCore::AttachNativeBrowser\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_BROWSER, IDH_CORE_CREATE_BROWSER);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_BROWSER, ATTACH_TO_BROWSER, IDH_CORE_CREATE_BROWSER);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


BOOL CCore::GetAppFilterData(const std::list<DescriptorToken>&tokens, String& sOutAppName, BOOL& bOutUseEqOp)
{
	for (std::list<DescriptorToken>::const_iterator it = tokens.begin();
	     it != tokens.end(); ++it)
	{
		const String& sAttributeName = it->m_sAttributeName;
		if (!_tcsicmp(sAttributeName.c_str(), BROWSER_APP_NAME))
		{
			sOutAppName = Common::ToUpper(it->m_sAttributePattern);
			bOutUseEqOp = (DescriptorToken::TOKEN_MATCH == it->m_operator);

			return TRUE;
		}
	}

	return FALSE;
}


BOOL CCore::IsValidDescriptorList(const std::list<DescriptorToken>& tokens)
{
	for (std::list<DescriptorToken>::const_iterator it = tokens.begin();
	     it != tokens.end(); ++it)
	{
		const String& sAttributeName = it->m_sAttributeName;
		if (_tcsicmp(sAttributeName.c_str(), BROWSER_TITLE_ATTR_NAME) &&
		    _tcsicmp(sAttributeName.c_str(), BROWSER_URL_ATTR_NAME)   &&
			_tcsicmp(sAttributeName.c_str(), BROWSER_APP_NAME))
		{
			return FALSE;
		}
	}

	return TRUE;
}


STDMETHODIMP CCore::Reset()
{
	FIRE_CANCEL_REQUEST();

	traceLog << "CCore::Reset\n";
	m_nSearchTimeout      = Settings::DEFAULT_SEARCH_TIMEOUT;
	m_nLoadTimeout        = Settings::DEFAULT_LOAD_TIMEOUT;
	m_nLastError          = Common::ERR_OK;
	m_vbLoadTimeoutIsErr  = Settings::DEFAULT_LOAD_TIMEOUT_IS_ERR;
	m_vbUseHardwareEvents = Settings::DEFAULT_USE_HARDWARE_EVENTS;

	return HRES_OK;
}


STDMETHODIMP CCore::get_IEVersion(BSTR* pBstrVersion)
{
	FIRE_CANCEL_REQUEST();

	// Reset lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pBstrVersion)
	{
		traceLog << "pBstrVersion is NULL in CCore::get_IEVersion\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, IE_VERSION_PROPERTY, IDH_CORE_IE_VERSION);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	CComBSTR bstrVersion;
	BOOL bRes = Common::GetIEVersion(bstrVersion);
	if (!bRes)
	{
		traceLog << "Common::GetIEVersion failed in CCore::get_IEVersion\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, IE_VERSION_PROPERTY, IDH_CORE_IE_VERSION);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	*pBstrVersion = bstrVersion.Detach();
	return HRES_OK;
}


BOOL CCore::FindVisibleChildShDocViewCallback(HWND hWnd, void*)
{
	// "Shell DocObject View" window must be a direct descendant of IEFrame (IE6) or "TabWindowClass" (IE7).
	// On IE6 (at least) there are IE bands that embeds IE controls but in this case the hierarchy has more levels until reaching "Shell DocObject View".
	ATLASSERT(Common::GetWndClass(hWnd) == _T("Shell DocObject View"));

	traceLog << "CCore::FindVisibleChildShDocViewCallback check window: " << hWnd << "\n";

	if (::IsWindow(hWnd) && ::IsWindowVisible(hWnd) && ::IsWindowEnabled(hWnd))
	{
		HWND hParentWnd = ::GetParent(hWnd);
		if (::IsWindow(hParentWnd))
		{
			int    nIeVer          = Common::GetIEVersion();
			String sParentWndClass = Common::GetWndClass(hParentWnd);

			if (nIeVer <= 6)
			{
				return (sParentWndClass == _T("IEFrame"));
			}
			else
			{
				return (sParentWndClass == _T("TabWindowClass"));
			}
		}
	}
	else
	{
		traceLog << "CCore::FindVisibleChildShDocViewCallback window: " << hWnd << "not visible/valid/enabled\n";
	}

	return FALSE;
}


BOOL CCore::FindVisibleChildWndCallback(HWND hWnd, void*)
{
	return (::IsWindow(hWnd) && ::IsWindowVisible(hWnd) && ::IsWindowEnabled(hWnd));
}


STDMETHODIMP CCore::get_foregroundBrowser(IBrowser** pFgBrowser)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pFgBrowser)
	{
		SetComErrorMessage(IDS_PROPERTY_FAILED, FOREGROUND_BROWSER_PROPERTY, IDH_CORE_FOREGROUND_BRWS);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	*pFgBrowser = NULL;

	HWND hFgWnd = ::GetForegroundWindow();
	BOOL bErr   = FALSE;

	if (NULL != hFgWnd)
	{
		if (Common::GetWndClass(hFgWnd) == _T("IEFrame"))
		{
			HWND hShellDocWnd = Common::GetChildWindowByClassName(hFgWnd, _T("Shell DocObject View"), FindVisibleChildShDocViewCallback);
			if (::IsWindow(hShellDocWnd))
			{
				ATLASSERT(Common::GetWndClass(hShellDocWnd) == _T("Shell DocObject View"));

				HWND hIeServerWnd = Common::GetChildWindowByClassName(hShellDocWnd, _T("Internet Explorer_Server"), FindVisibleChildWndCallback);
				if (::IsWindow(hIeServerWnd))
				{
					CComQIPtr<IExplorerPlugin> spPlugin = NewMarshalService::FindBrwsPluginByIeWnd(hIeServerWnd);
					if (spPlugin != NULL)
					{
						IBrowser* pBrws = NULL;
						HRESULT   hRes  = CComCoClass<CBrowser>::CreateInstance(&pBrws);
						if (pBrws != NULL)
						{
							ATLASSERT(SUCCEEDED(hRes));

							CBrowser* pBrowser = static_cast<CBrowser*>(pBrws);	// Down cast !!!
							ATLASSERT(pBrowser != NULL);
							pBrowser->SetPlugin(spPlugin);
							pBrowser->SetCore(this);

							*pFgBrowser = pBrowser;
						}
						else
						{
							traceLog << "Can not create Browser instance in CCore::get_foregroundBrowser\n";
							bErr = TRUE;
						}
					}
					else
					{
						traceLog << "Can NOT find IExplorerPlugin in CCore::get_foregroundBrowser\n";
						bErr = TRUE;
					}
				}
				else
				{
					traceLog << "Can NOT find 'Internet Explorer_Server' window in CCore::get_foregroundBrowser\n";
					bErr = TRUE;
				}
			}
			else
			{
				traceLog << "Can NOT find 'Shell DocObject View' window in CCore::get_foregroundBrowser\n";
				bErr = TRUE;
			}
		}
	}

	if (bErr)
	{
		SetComErrorMessage(IDS_PROPERTY_FAILED, FOREGROUND_BROWSER_PROPERTY, IDH_CORE_FOREGROUND_BRWS);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CCore::get_productVersion(BSTR* pBstrVersion)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pBstrVersion)
	{
		SetComErrorMessage(IDS_PROPERTY_FAILED, PRODUCT_VERSION_CORE_PROPERTY, IDH_CORE_PROD_VERSION);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	// No revision number. It will alwasy be zero.
	Common::Ostringstream outputStream;
	outputStream << (int)ProductSettings::MAJOR_VERSION << _T(".") << (int)ProductSettings::MINOR_VERSION << _T(".") << (int)ProductSettings::BUILD_NUMBER << _T(".0");

	CComBSTR bstrVersion = outputStream.str().c_str();
	*pBstrVersion = bstrVersion.Detach();

	return HRES_OK;
}


STDMETHODIMP CCore::get_productName(BSTR* pBstrName)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pBstrName)
	{
		SetComErrorMessage(IDS_PROPERTY_FAILED, PRODUCT_NAME_CORE_PROPERTY, IDH_CORE_PROD_NAME);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	CComBSTR bstrProductName = ProductSettings::OPEN_TWEBST_PRODUCT_NAME;
	*pBstrName = bstrProductName.Detach();

	return HRES_OK;
}


STDMETHODIMP CCore::GetClipboardText(BSTR* pBstrClipboardText)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pBstrClipboardText)
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, GET_CLPBRD_TEXT_CORE_METHOD, IDH_CORE_GET_CLIPBOARD_TEXT);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	CComBSTR bstrResult = L"";

	if (::OpenClipboard(NULL))
	{
		if (::IsClipboardFormatAvailable(CF_TEXT) || ::IsClipboardFormatAvailable(CF_UNICODETEXT))
		{
			BOOL   bUnicode   = TRUE;
			HANDLE hClipboard = ::GetClipboardData(CF_UNICODETEXT);

			if (NULL == hClipboard)
			{
				bUnicode   = FALSE;
				hClipboard = ::GetClipboardData(CF_TEXT);
			}

			if (hClipboard != NULL)
			{
				LPCSTR szClipboardData = (LPCSTR)::GlobalLock(hClipboard);

				if (szClipboardData != NULL)
				{
					if (bUnicode)
					{
						LPCWSTR szClipboardWText = (LPCWSTR)szClipboardData;

						bstrResult       = szClipboardWText;
						szClipboardWText = NULL;
					}
					else
					{
						LPCSTR szClipboardText = (LPCSTR)szClipboardData;

						bstrResult      = szClipboardText;
						szClipboardText = NULL;
					}

					::GlobalUnlock(hClipboard);

					*pBstrClipboardText = bstrResult.Detach();
					hClipboard = NULL;
				}
				else
				{
					traceLog << "GlobalLock failed in CCore::GetClipboardText\n";
				}
			}
			else
			{
				traceLog << "GetClipboardData failed in CCore::GetClipboardText\n";
			}

			hClipboard = NULL;
		}
		else
		{
			traceLog << "CF_TEXT NOT available in CCore::GetClipboardText\n";
		}

		BOOL bRes = ::CloseClipboard();
		ATLASSERT(bRes);
	}
	else
	{
		traceLog << "Can NOT OpenClipboard in CCore::GetClipboardText\n";
	}

	return HRES_OK;
}


STDMETHODIMP CCore::SetClipboardText(BSTR bstrClipboardText)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == bstrClipboardText)
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, SET_CLPBRD_TEXT_CORE_METHOD, IDH_CORE_SET_CLIPBOARD_TEXT);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	// Empty the clipboard.
	HRESULT hRes = ClearClipboard();

	if (FAILED(hRes))
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, SET_CLPBRD_TEXT_CORE_METHOD, IDH_CORE_SET_CLIPBOARD_TEXT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	BOOL bErr       = FALSE;

	// Put text into clipboard.
	if (::OpenClipboard(NULL))
	{
		// Put unicode text in clipboard.
		if (!SetTextInClipboard(bstrClipboardText))
		{
			bErr = TRUE;
		}

		// Put ansi text in clipboard.
		USES_CONVERSION;
		if (!SetTextInClipboard(W2A(bstrClipboardText)))
		{
			bErr = TRUE;
		}

		BOOL bRes = ::CloseClipboard();
		ATLASSERT(bRes);
	}
	else
	{
		bErr = TRUE;
		traceLog << "Can NOT OpenClipboard in CCore::SetClipboardText\n";
	}

	if (bErr)
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, SET_CLPBRD_TEXT_CORE_METHOD, IDH_CORE_SET_CLIPBOARD_TEXT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	return HRES_OK;
}


BOOL CCore::SetTextInClipboard(LPCSTR szText)
{
	ATLASSERT(szText != NULL);

	size_t  nCharCount = strlen(szText);
	BOOL    bErr       = FALSE;
	HGLOBAL hglb       = ::GlobalAlloc(GMEM_MOVEABLE, (nCharCount + 1) * sizeof(CHAR));

	if (hglb != NULL)
	{
		LPSTR szBuff = (LPSTR)::GlobalLock(hglb);

		strncpy_s(szBuff, nCharCount + 1, szText, nCharCount);
		szBuff[nCharCount] = L'\0';

		::GlobalUnlock(hglb);

		HANDLE hData = SetClipboardData(CF_TEXT, hglb);
		if (NULL == hData)
		{
			::GlobalFree(hglb);
			hglb = NULL;
			bErr = TRUE;

			traceLog << "SetClipboardData failed in CCore::SetTextInClipboard\n";
		}
	}
	else
	{
		bErr = TRUE;
		traceLog << "GlobalAlloc failed in CCore::SetTextInClipboard\n";
	}

	return !bErr;
}


BOOL CCore::SetTextInClipboard(LPCWSTR szText)
{
	ATLASSERT(szText != NULL);

	size_t  nCharCount = wcslen(szText);
	BOOL    bErr       = FALSE;
	HGLOBAL hglb       = ::GlobalAlloc(GMEM_MOVEABLE, (nCharCount + 1) * sizeof(WCHAR));

	if (hglb != NULL)
	{
		LPWSTR szBuff = (LPWSTR)::GlobalLock(hglb);

		wcsncpy_s(szBuff, nCharCount + 1, szText, nCharCount);
		szBuff[nCharCount] = L'\0';

		::GlobalUnlock(hglb);

		HANDLE hData = SetClipboardData(CF_UNICODETEXT, hglb);
		if (NULL == hData)
		{
			::GlobalFree(hglb);
			hglb = NULL;
			bErr = TRUE;

			traceLog << "SetClipboardData failed in CCore::SetTextInClipboard\n";
		}
	}
	else
	{
		bErr = TRUE;
		traceLog << "GlobalAlloc failed in CCore::SetTextInClipboard\n";
	}

	return !bErr;
}


HRESULT CCore::ClearClipboard()
{
	HRESULT hRes = S_OK;

	if (::OpenClipboard(NULL))
	{
		BOOL bRes = ::EmptyClipboard();
		if (!bRes)
		{
			traceLog << "Can NOT EmptyClipboard in CCore::ClearClipboard\n";
			hRes = E_FAIL;
		}

		bRes = ::CloseClipboard();
		if (!bRes)
		{
			traceLog << "Can NOT CloseClipboard in CCore::ClearClipboard\n";
			hRes = E_FAIL;
		}
	}
	else
	{
		traceLog << "Can NOT OpenClipboard in CCore::ClearClipboard\n";
		hRes = E_FAIL;
	}

	return hRes;
}


void CCore::FindAllIEFramesWnd(std::set<HWND>& allIeFrames)
{
	ATLASSERT(allIeFrames.empty());
	ATLASSERT((Common::GetIEVersion() > 7) && !Common::IsWindowsVistaOrLater());

	// Enum top level windows seraching a "IEFrame" window that belongs to dwTid thread.
	::EnumWindows(FindAllIEFramesWndCallback, (LPARAM)(&allIeFrames));
}


BOOL CALLBACK CCore::FindAllIEFramesWndCallback(HWND hWnd, LPARAM lParam)
{
	if (GetWndClass(hWnd) == _T("IEFrame"))
	{
		std::set<HWND>* pIEFramesSet = reinterpret_cast<std::set<HWND>*>(lParam);
		ATLASSERT(pIEFramesSet != NULL);

		pIEFramesSet->insert(hWnd);
	}

	return TRUE;
}


HWND CCore::FindIEFrameExcludeSet(std::set<HWND>& ieFramesToExclude)
{
	ATLASSERT((Common::GetIEVersion() > 7) && !Common::IsWindowsVistaOrLater());

	FindIEFrameExcludeSetInfo ieFrameInfo(&ieFramesToExclude);

	DWORD dwStartTime = ::GetTickCount();
	while(TRUE)
	{
		::EnumWindows(FindIEFrameExcludeSetCallback, (LPARAM)(&ieFrameInfo));
		if (ieFrameInfo.m_hIEFrameWnd != NULL)
		{
			// We have found the window.
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		ProcessMessagesOrSleep();
	}

	if (Common::GetIEVersion() > 6)
	{
		// IE7 or IE8 on XP.
		if (ieFrameInfo.m_hIEFrameWnd != NULL)
		{
			while (TRUE)
			{
				// Find a tab.
				HWND hTabWnd = FindTabWnd(ieFrameInfo.m_hIEFrameWnd);
				if (hTabWnd != NULL)
				{
					// Tab window found.
					return hTabWnd;
				}

				DWORD dwCurrentTime = ::GetTickCount();
				if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
				{
					// Timeout has expired. Quit the loop.
					break;
				}

				ProcessMessagesOrSleep();
			}
		}

		return NULL;
	}
	else
	{
		return ieFrameInfo.m_hIEFrameWnd;
	}
}


BOOL CALLBACK CCore::FindIEFrameExcludeSetCallback(HWND hWnd, LPARAM lParam)
{
	if (GetWndClass(hWnd) == _T("IEFrame"))
	{
		FindIEFrameExcludeSetInfo* pIEFrameInfo = reinterpret_cast<FindIEFrameExcludeSetInfo*>(lParam);
		ATLASSERT(pIEFrameInfo != NULL);

		if (pIEFrameInfo->m_pIEFramesToExclude->find(hWnd) == pIEFrameInfo->m_pIEFramesToExclude->end())
		{
			pIEFrameInfo->m_hIEFrameWnd = hWnd;
			return FALSE;
		}
	}

	return TRUE;
}


STDMETHODIMP CCore::get_closeBrowserPopups(VARIANT_BOOL* pAutoClosePopups)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pAutoClosePopups)
	{
		traceLog << "Invalid argument in  CCore::get_closeBrowserPopups\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_CLOSE_BROWSER_POPUPS, IDH_CORE_CLOSE_BROWSER_POPUPS);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	*pAutoClosePopups = m_vbAutoClosePopups;
	return HRES_OK;
}


STDMETHODIMP CCore::put_closeBrowserPopups(VARIANT_BOOL  vbAutoClosePopups)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	m_vbAutoClosePopups = vbAutoClosePopups;
	return HRES_OK;
}


BOOL CCore::GetAutoClosePopups()
{
	return (VARIANT_TRUE == m_vbAutoClosePopups);
}


STDMETHODIMP CCore::get_asyncHtmlEvents(VARIANT_BOOL* pVal)
{
	FIRE_CANCEL_REQUEST();

	if (NULL == pVal)
	{
		traceLog << "Invalid argument in  CCore::get_asyncHtmlEvents\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_ASYNC_HTML_EVENTS, IDH_CORE_ASYNC_HTML_EVENTS);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	*pVal = m_bAsyncEvents;
	return HRES_OK;
}


STDMETHODIMP CCore::put_asyncHtmlEvents(VARIANT_BOOL newVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);
	m_bAsyncEvents = newVal;

	return HRES_OK;
}


void CCore::ProcessMessagesOrSleep()
{
/*
	// In WinRunner we need to process messages otherwise the newly started IE is locked until timeout is reached.
	// maybe WinRunner installs some hooks and we need to process messages.

	MSG	msg    = { 0 };
	int nCount = 0;

	// Dispatch 256 messages at most.
	while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE) && (nCount < 256))
	{
		::DispatchMessage(&msg);
		nCount++;
	}

	::Sleep(Common::INTERNAL_GLOBAL_PAUSE);
*/

	::Sleep(Common::INTERNAL_GLOBAL_PAUSE);
}


STDMETHODIMP CCore::FindElementFromPoint(LONG x, LONG y, IElement** ppElement)
{
	if (!ppElement)
	{
		traceLog << "Invalid arg in CCore::FindElementFromPoint\n";

		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, CORE_FIND_ELEM_FROM_POINT, IDH_CORE_ELEMENT_FROM_POINT);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	HWND hIEWnd = HtmlHelpers::GetIEServerFromScrPt(x, y);
	if (NULL == hIEWnd)
	{
		traceLog << "No 'Internet Explorer_Server' window found in  CCore::FindElementFromPoint\n";
		return HRES_OK;
	}

	ATLASSERT(Common::GetWndClass(hIEWnd) == _T("Internet Explorer_Server"));

	CComQIPtr<IExplorerPlugin> spPlugin = NewMarshalService::FindBrwsPluginByIeWnd(hIEWnd);
	if (spPlugin == NULL)
	{
		traceLog << "FindBrwsPluginByIeWnd returned NULL in CCore::FindElementFromPoint\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, CORE_FIND_ELEM_FROM_POINT, IDH_CORE_ELEMENT_FROM_POINT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	CComQIPtr<IHTMLElement> spElem;
	HRESULT hRes = spPlugin->FindElementFromPoint(x, y, &spElem);
	if (FAILED(hRes) || !spElem)
	{
		traceLog << "spPlugin->FindElementFromPoint failed in CCore::FindElementFromPoint\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, CORE_FIND_ELEM_FROM_POINT, IDH_CORE_ELEMENT_FROM_POINT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	// Create a new element object.
	IElement* pNewElement = NULL;
	hRes = CComCoClass<CElement>::CreateInstance(&pNewElement);

	if (pNewElement != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CElement* pElementObject = static_cast<CElement*>(pNewElement); // Down cast !!!
		ATLASSERT(pElementObject != NULL);

		pElementObject->SetPlugin(spPlugin);
		pElementObject->SetCore(this);
		pElementObject->SetHtmlElement(spElem);

		// Set the object reference into the output pointer.
		*ppElement = pNewElement;
		return HRES_OK;
	}
	else
	{
		traceLog << "Can not create a Element object in CCore::AttachToNativeElement\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_CORE_ELEMENT_FROM_POINT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}
}


STDMETHODIMP CCore::AttachToHWND(LONG nWnd, IBrowser** ppBrowser)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	HWND hTargetWnd = (HWND)LongToHandle(nWnd);
	if (!::IsWindow(hTargetWnd))
	{
		traceLog << "Invalid window handle param in CCore::AttachToWnd\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_WND_CORE_METHOD, IDH_CORE_ATTACH_TO_WND);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	if (NULL == ppBrowser)
	{
		traceLog << "Invalid out param pointer in CCore::AttachToWnd\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_WND_CORE_METHOD, IDH_CORE_ATTACH_TO_WND);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	String sWndClassName = Common::GetWndClass(hTargetWnd);
	BOOL   bRes          = TRUE;

	if (_T("IEFrame") == sWndClassName)
	{
		bRes = AttachToIEFrame(hTargetWnd, ppBrowser);
	}
	else if (_T("TabWindowClass") == sWndClassName)
	{
		if (Common::GetIEVersion() <= 6)
		{
			traceLog << "CCore::AttachToWnd(TabWindowClass) only works for IE7\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, ATTACH_TO_WND_CORE_METHOD, IDH_CORE_ATTACH_TO_WND);
			SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);

			return HRES_OPERATION_NOT_APPLICABLE;
		}

		bRes = AttachToTab(hTargetWnd, ppBrowser);
	}
	else if (_T("Internet Explorer_Server") == sWndClassName)
	{
		bRes = AttachToIEServer(hTargetWnd, ppBrowser);
	}
	/*else if (_T("Internet Explorer_TridentDlgFrame") == sWndClassName)
	{
		// Future implementation to support HTML dialogs.
		bRes = AttachToTridentDlg(hTargetWnd, ppBrowser);
	}*/
	else
	{
		HWND hIeServerWnd = Common::GetChildWindowByClassName(hTargetWnd, _T("Internet Explorer_Server"), FindVisibleChildWndCallback);
		if (::IsWindow(hIeServerWnd))
		{
			return AttachToIEServer(hIeServerWnd, ppBrowser);
		}

		traceLog << "Invalid window type param in CCore::AttachToWnd\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, ATTACH_TO_WND_CORE_METHOD, IDH_CORE_ATTACH_TO_WND);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	if (!bRes)
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ATTACH_TO_WND_CORE_METHOD, IDH_CORE_ATTACH_TO_WND);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	return HRES_OK;
}


BOOL CCore::AttachToIEFrame(HWND hIEFrameWnd, IBrowser** ppBrowser)
{
	ATLASSERT(::IsWindow(hIEFrameWnd));
	ATLASSERT(_T("IEFrame") == Common::GetWndClass(hIEFrameWnd));
	ATLASSERT(ppBrowser != NULL);

	*ppBrowser = NULL;

	HWND hShellDocWnd = Common::GetChildWindowByClassName(hIEFrameWnd, _T("Shell DocObject View"), FindVisibleChildShDocViewCallback);
	if (::IsWindow(hShellDocWnd))
	{
		ATLASSERT(Common::GetWndClass(hShellDocWnd) == _T("Shell DocObject View"));

		HWND hIeServerWnd = Common::GetChildWindowByClassName(hShellDocWnd, _T("Internet Explorer_Server"), FindVisibleChildWndCallback);
		if (::IsWindow(hIeServerWnd))
		{
			return AttachToIEServer(hIeServerWnd, ppBrowser);
		}
		else
		{
			traceLog << "Can NOT find 'Internet Explorer_Server' window in CCore::AttachToIEFrame\n";
			return FALSE;
		}
	}
	else
	{
		traceLog << "Can NOT find 'Shell DocObject View' window in CCore::AttachToIEFrame\n";
		return FALSE;
	}

	return TRUE;
}


BOOL CCore::AttachToTab(HWND hTabWnd, IBrowser** ppBrowser)
{
	ATLASSERT(::IsWindow(hTabWnd));
	ATLASSERT(_T("TabWindowClass") == Common::GetWndClass(hTabWnd));
	ATLASSERT(ppBrowser != NULL);
	ATLASSERT(Common::GetIEVersion() > 6);

	*ppBrowser = NULL;

	HWND hShellDocWnd = Common::GetChildWindowByClassName(hTabWnd, _T("Shell DocObject View"), NewMarshalService::FindChildShDocViewCallback);
	if (::IsWindow(hShellDocWnd))
	{
		ATLASSERT(Common::GetWndClass(hShellDocWnd) == _T("Shell DocObject View"));

		HWND hIeServerWnd = Common::GetChildWindowByClassName(hShellDocWnd, _T("Internet Explorer_Server"));
		if (::IsWindow(hIeServerWnd))
		{
			return AttachToIEServer(hIeServerWnd, ppBrowser);
		}
		else
		{
			traceLog << "Can NOT find 'Internet Explorer_Server' window in CCore::AttachToTab\n";
			return FALSE;
		}
	}
	else
	{
		traceLog << "Can NOT find 'Shell DocObject View' window in CCore::AttachToTab\n";
		return FALSE;
	}

	return TRUE;
}


// This will only work with windows coming from IE or hosted browsers already attached!!!
BOOL CCore::AttachToIEServer(HWND hIeServerWnd, IBrowser** ppBrowser)
{
	ATLASSERT(::IsWindow(hIeServerWnd));
	ATLASSERT(_T("Internet Explorer_Server") == Common::GetWndClass(hIeServerWnd));
	ATLASSERT(ppBrowser != NULL);

	*ppBrowser = NULL;

	CComQIPtr<IExplorerPlugin> spPlugin = NewMarshalService::FindBrwsPluginByIeWnd(hIeServerWnd, 0, TRUE);
	if (spPlugin != NULL)
	{
		IBrowser* pBrws = NULL;
		HRESULT   hRes  = CComCoClass<CBrowser>::CreateInstance(&pBrws);
		if (pBrws != NULL)
		{
			ATLASSERT(SUCCEEDED(hRes));

			CBrowser* pBrowser = static_cast<CBrowser*>(pBrws);
			ATLASSERT(pBrowser != NULL);
			pBrowser->SetPlugin(spPlugin);
			pBrowser->SetCore(this);

			*ppBrowser = pBrowser;
		}
		else
		{
			traceLog << "Can not create Browser instance in CCore::AttachToIEServer\n";
			return FALSE;
		}
	}
	else
	{
		traceLog << "Can NOT find IExplorerPlugin in CCore::AttachToIEServer\n";
		return FALSE;
	}

	return TRUE;
}


BOOL CCore::AttachToTridentDlg(HWND hTargetWnd, IBrowser** ppBrowser)
{
	ATLASSERT(::IsWindow(hTargetWnd));
	ATLASSERT(_T("Internet Explorer_TridentDlgFrame") == Common::GetWndClass(hTargetWnd));
	ATLASSERT(ppBrowser != NULL);

	*ppBrowser = NULL;

	return FALSE;
}
