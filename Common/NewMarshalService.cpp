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
#include "..\OTWBSTInjector\OTWBSTInjector.h"
#include "Common.h"
#include "DebugServices.h"
#include "MarshalService.h"



namespace NewMarshalService
{
	struct FindIeServerListData
	{
		FindIeServerListData(std::list<HWND>* pResultList, const String& sAppName, BOOL bUseEqualOp) : 
		        m_pResultList(pResultList), m_sAppName(Common::ToUpper(sAppName)), m_bUseEqualOp(bUseEqualOp)
		{
			ATLASSERT(pResultList != NULL);

			if (_T("") == m_sAppName)
			{
				m_sAppName = _T("IEXPLORE.EXE");
			}
		}

		BOOL TopWndPassTest(HWND hTopWnd)
		{
			if (m_sAppName == _T("*"))
			{
				return m_bUseEqualOp;
			}
			else if (m_sAppName == _T("IEXPLORE.EXE"))
			{
				// Only iexplore.exe
				return (Common::GetWndClass(hTopWnd) == _T("IEFrame"));
			}
			else
			{
				String sProcName;
				if (!Common::GetWndProcName(hTopWnd, sProcName))
				{
					traceLog << "Cannot get process name in FindIeServerListData::TopWndPassTest for hwnd=" << hTopWnd <<"\n";
					return FALSE;
				}

				sProcName = Common::ToUpper(sProcName);
				return (m_sAppName == sProcName); // Wildcard match needed? I guess not ...
			}
		}

		BOOL IePassTest(HWND hIeWnd)
		{
			ATLASSERT(_T("Internet Explorer_Server") == Common::GetWndClass(hIeWnd));

			if (m_sAppName == _T("*"))
			{
				return m_bUseEqualOp;
			}
			else
			{
				String sProcName;
				if (!Common::GetWndProcName(hIeWnd, sProcName))
				{
					traceLog << "Cannot get process name in FindIeServerListData::IePassTest for hwnd=" << hIeWnd <<"\n";
					return FALSE;
				}

				sProcName = Common::ToUpper(sProcName);
				return (m_sAppName == sProcName);
			}
		}

		String           m_sAppName;
		BOOL             m_bUseEqualOp;
		std::list<HWND>* m_pResultList;
	};


	struct FindNewTabData
	{
		FindNewTabData(std::list<HWND>* pIeTabList) : m_pTabList(pIeTabList)
		{
			ATLASSERT(pIeTabList != NULL);
			m_hResultWnd = NULL;
		}

		std::list<HWND>* m_pTabList;
		HWND             m_hResultWnd;
	};


	HWND          GetIEServerNoTimeout (HWND hParentWnd);
	BOOL          IsTabIEServer        (HWND hIeServer);
	HWND          GetIEServerWnd       (HWND hParentWnd, DWORD nTimeout);
	BOOL          IsIEWindow           (HWND hIeWnd);
	BOOL CALLBACK EnumTopWndCallback   (HWND hWnd, LPARAM lParam);
	BOOL CALLBACK EnumIeWndCallback    (HWND hWnd, LPARAM lParam);
	BOOL CALLBACK EnumTabsWndCallback  (HWND hWnd, LPARAM lParam);
	BOOL CALLBACK FindNewTabWndCallback(HWND hWnd, LPARAM lParam);
}


CComQIPtr<IExplorerPlugin> NewMarshalService::FindBrwsPluginByIeWnd
	(
		HWND                    hIeServer,
		DWORD                   nTimeout,
		BOOL                    bStopBrowser,
		CComQIPtr<IWebBrowser2> spBrowser
	)
{
	if (!::IsWindow(hIeServer))
	{
		traceLog << "Invalid hIeServer in NewMarshalService::FindBrwsPluginByIeWnd\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	CComQIPtr<IWebBrowser2> spBrowserToAttach;
	BOOL bSameThread = (::GetCurrentThreadId() == ::GetWindowThreadProcessId(hIeServer, NULL));

	if (bSameThread)
	{
		spBrowserToAttach = spBrowser;
	}

	HWND hIeServerWnd = hIeServer;
	if (!spBrowserToAttach)
	{
		// There is no IWebBrowser2 in the same thread to attach to. Use "Internet Explorer_Server" window to inject.
		hIeServerWnd = GetIEServerWnd(hIeServer, nTimeout);
		if (NULL == hIeServerWnd)
		{
			traceLog << "Invalid param window in NewMarshalService::FindBrwsPluginByIeWnd\n";
			return CComQIPtr<IExplorerPlugin>();
		}
	}

	ATLASSERT((_T("Internet Explorer_Server") == Common::GetWndClass(hIeServerWnd)) || (spBrowserToAttach != NULL));

	if (!InjectBHO(hIeServerWnd, bStopBrowser, spBrowserToAttach))
	{
		traceLog << "Can NOT inject BHO in NewMarshalService::FindBrwsPluginByIeWnd\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	String sPluginMarshalName;
	BOOL bName = GetMarshalWndName(hIeServerWnd, sPluginMarshalName);
	if (!bName)
	{
		traceLog << "GetMarshalWndName failed in NewMarshalService::FindBrwsPluginByIeWnd\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	HWND hPluginCommWnd = ::FindWindow(HIDDEN_COMMUNICATION_WND_CLASS_NAME, sPluginMarshalName.c_str());
	if (NULL == hPluginCommWnd)
	{
		traceLog << "Can NOT find communication window in NewMarshalService::FindBrwsPluginByIeWnd\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	CComQIPtr<IAccessible> spAccPlugin;
	HRESULT hRes = AccessibleObjectFromWindow(hPluginCommWnd, (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)&spAccPlugin);

	if (FAILED(hRes) || (spAccPlugin == NULL))
	{
		traceLog << "Can NOT get IAccessible object from wnd in NewMarshalService::FindBrwsPluginByIeWnd\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	CComQIPtr<IServiceProvider> spServProv = spAccPlugin;
	CComQIPtr<IExplorerPlugin>  spPlugin;

	hRes = spServProv->QueryService(IID_IExplorerPlugin, IID_IExplorerPlugin, (void**)&spPlugin);
	return spPlugin;
}


BOOL NewMarshalService::GetMarshalWndName(HWND hIeServer, String& sRetName)
{
	DWORD dwIeServerThID = ::GetWindowThreadProcessId(hIeServer, NULL);
	BOOL  bIsSameThread  = (::GetCurrentThreadId() == dwIeServerThID);
	BOOL  bIsIeWindow    = IsIEWindow(hIeServer);

	if (!bIsIeWindow && !bIsSameThread)
	{
		traceLog << "Invalid param window in NewMarshalService::GetMarshalWndName\n";
		return FALSE;
	}
	else
	{
		const TCHAR PLUGIN_PREFIX_NAME[] = _T("{67850823-5AF3-4188-81F0-3FC00AF78281}+");
		Ostringstream os;

		if (!bIsSameThread)
		{
			os << (ULONG)HandleToLong(hIeServer);
		}
		else
		{
			if (!bIsIeWindow)
			{
				os << L"+" << dwIeServerThID;
			}
			else
			{
				os << (ULONG)HandleToLong(hIeServer);
			}
		}

		sRetName = PLUGIN_PREFIX_NAME + os.str();
		return TRUE;
	}
}


void NewMarshalService::GetIEServerList(std::list<HWND>& ieServerList, const String& sAppName, BOOL bEqual)
{
	ieServerList.clear();

	FindIeServerListData findIeData(&ieServerList, sAppName, bEqual);
	::EnumWindows(EnumTopWndCallback, (LPARAM)(&findIeData));
}


/*void NewMarshalService::GetIEServerList(std::list<HWND>& ieServerList, BOOL bIeOnly)
{
	ieServerList.clear();

	FindIeServerListData findIeData(&ieServerList, bIeOnly);
	::EnumWindows(EnumTopWndCallback, (LPARAM)(&findIeData));
}*/


void NewMarshalService::GetTabsIEServer
	(
		HWND             hIEFrame,
		std::list<HWND>& ieServerList
	)
{
	if (GetWndClass(hIEFrame) != _T("IEFrame"))
	{
		return;
	}

	ieServerList.clear();
	::EnumChildWindows(hIEFrame, EnumTabsWndCallback, (LPARAM)(&ieServerList));
}


BOOL NewMarshalService::FindIeWndNotInList
	(
		HWND             hIEFrame,
		std::list<HWND>& ieServerList,
		HWND&            hNewThreadWnd,
		DWORD            nTimeout
	)
{
	if (GetWndClass(hIEFrame) != _T("IEFrame"))
	{
		return FALSE;
	}

	DWORD dwStartTime = ::GetTickCount();
	while (TRUE)
	{
		FindNewTabData findNewTabInfo(&ieServerList);
		::EnumWindows(FindNewTabWndCallback, (LPARAM)(&findNewTabInfo));

		hNewThreadWnd = findNewTabInfo.m_hResultWnd;
		if (hNewThreadWnd != NULL)
		{
			// Tab found!
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > nTimeout)
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		::Sleep(Common::INTERNAL_GLOBAL_PAUSE);
	}

	return TRUE;
}


HWND NewMarshalService::GetIEServerWnd(HWND hParentWnd, DWORD nTimeout)
{
	if (!::IsWindow(hParentWnd))
	{
		traceLog << "Invalid param window in NewMarshalService::GetIEServerWnd\n";
		return NULL;
	}

	HWND  hResultWnd  = NULL;
	DWORD dwStartTime = ::GetTickCount();

	while (TRUE)
	{
		hResultWnd = GetIEServerNoTimeout(hParentWnd);
		if (hResultWnd != NULL)
		{
			// Window found!
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > nTimeout)
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		::Sleep(Common::INTERNAL_GLOBAL_PAUSE);
	}

	return hResultWnd;
}


HWND NewMarshalService::GetIEServerNoTimeout(HWND hParentWnd)
{
	ATLASSERT(::IsWindow(hParentWnd));

	String parentClass = Common::GetWndClass(hParentWnd);
	if (_T("Internet Explorer_Server") == parentClass)
	{
		traceLog << "Internet Explorer_Server case in NewMarshalService::GetIEServerWnd\n";
		return hParentWnd;
	}
	else if (_T("IEFrame") == parentClass)
	{
		ATLASSERT((Common::GetIEVersion() == 6) || ((Common::GetIEVersion() == 8) && !Common::IsWindowsVistaOrLater()));
		traceLog << "IEFrame case in NewMarshalService::GetIEServerWnd\n";

		HWND hShellDocWnd = Common::GetChildWindowByClassName(hParentWnd, _T("Shell DocObject View"), FindChildShDocViewCallback);
		if (::IsWindow(hShellDocWnd))
		{
			ATLASSERT(Common::GetWndClass(hShellDocWnd) == _T("Shell DocObject View"));

			HWND hIeServerWnd = Common::GetChildWindowByClassName(hShellDocWnd, _T("Internet Explorer_Server"));
			return hIeServerWnd;
		}
		else
		{
			traceLog << "Cannot get Shell DocObject View in NewMarshalService::GetIEServerWnd\n";
			return NULL;
		}
	}
	else if (_T("TabWindowClass") == parentClass)
	{
		ATLASSERT(Common::GetIEVersion() > 6);

		traceLog << "TabWindowClass case in NewMarshalService::GetIEServerWnd\n";
		return Common::GetChildWindowByClassName(hParentWnd, _T("Internet Explorer_Server"));
	}
	else
	{
		return NULL;
	}
}


BOOL NewMarshalService::IsIEWindow(HWND hIeWnd)
{
	if (!::IsWindow(hIeWnd))
	{
		return FALSE;
	}
	else
	{
		return (Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));
	}
}


BOOL NewMarshalService::FindChildShDocViewCallback(HWND hWnd, void*)
{
	// "Shell DocObject View" window must be a direct descendant of IEFrame (IE6) or "TabWindowClass" (IE7).
	// On IE6 (at least) there are IE bands that embeds IE controls but in this case the hierarchy has more levels until reaching "Shell DocObject View".
	ATLASSERT(Common::GetWndClass(hWnd) == _T("Shell DocObject View"));

	if (::IsWindow(hWnd))
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

	return FALSE;
}


BOOL CALLBACK NewMarshalService::EnumTopWndCallback(HWND hWnd, LPARAM lParam)
{
	FindIeServerListData* pData = (FindIeServerListData*)lParam;
	if (!pData || !::IsWindow(hWnd))
	{
		// Error, stop enumerating.
		return FALSE;
	}

	if (!pData->TopWndPassTest(hWnd))
	{
		// Not a IE window, skip it.
		return TRUE;
	}

	::EnumChildWindows(hWnd, EnumIeWndCallback, lParam);
	return TRUE;
}


BOOL CALLBACK NewMarshalService::EnumIeWndCallback(HWND hWnd, LPARAM lParam)
{
	FindIeServerListData* pData = (FindIeServerListData*)lParam;
	if (!pData || !::IsWindow(hWnd))
	{
		// Error, stop enumerating.
		return FALSE;
	}

	if (Common::GetWndClass(hWnd) == _T("Internet Explorer_Server"))
	{
		if (pData->IePassTest(hWnd))
		{
			pData->m_pResultList->push_back(hWnd);
		}
	}

	return TRUE;
}


BOOL CALLBACK NewMarshalService::EnumTabsWndCallback(HWND hWnd, LPARAM lParam)
{
	if (!lParam || !::IsWindow(hWnd))
	{
		// Error, stop enumerating.
		return FALSE;
	}

	if (IsTabIEServer(hWnd))
	{
		std::list<HWND>* pIeList = (std::list<HWND>*)lParam;
		pIeList->push_back(hWnd);
	}

	return TRUE;
}


BOOL CALLBACK NewMarshalService::FindNewTabWndCallback(HWND hWnd, LPARAM lParam)
{
	if (!lParam || !::IsWindow(hWnd))
	{
		// Error, stop enumerating.
		return FALSE;
	}

	if (IsTabIEServer(hWnd))
	{
		FindNewTabData* pData = (FindNewTabData*)lParam;
		std::list<HWND>::const_iterator it = std::find(pData->m_pTabList->begin(), pData->m_pTabList->end(), hWnd);
		if (it == pData->m_pTabList->end())
		{
			pData->m_hResultWnd = hWnd;
			return FALSE;
		}
	}

	return TRUE;
}


BOOL NewMarshalService::IsTabIEServer(HWND hIeServer)
{
	if (Common::GetWndClass(hIeServer) == _T("Internet Explorer_Server"))
	{
		HWND hShellDocWnd = ::GetParent(hIeServer);
		if (Common::GetWndClass(hShellDocWnd) == _T("Shell DocObject View"))
		{
			HWND hTabWnd = ::GetParent(hShellDocWnd);
			if (Common::GetWndClass(hTabWnd) == _T("TabWindowClass"))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}
