#include "StdAfx.h"
#include "DebugServices.h"
#include "MarshalService.h"



namespace MarshalService
{
	// Local functions.
	BOOL CALLBACK FindMarshallWndCallback      (HWND hWnd, LPARAM lParam);
	BOOL CALLBACK FindTabMarshallWndCallback   (HWND hWnd, LPARAM lParam);
	BOOL CALLBACK FindNewTabMarshallWndCallback(HWND hWnd, LPARAM lParam);

	struct FIND_TABS_INFO
	{
		FIND_TABS_INFO(HWND hIEFrame, std::list<DWORD>* pTabThreadList) : 
			m_hIEFrame(hIEFrame), m_pTabThreadList(pTabThreadList)
		{
			ATLASSERT(GetWndClass(hIEFrame) == _T("IEFrame"));
			ATLASSERT(pTabThreadList       != NULL);

			m_dwThreadResult = 0;
		}

		HWND              m_hIEFrame;
		std::list<DWORD>* m_pTabThreadList;
		DWORD             m_dwThreadResult;
	};
}



String MarshalService::GetMarshalWndName(DWORD dwThreadID)
{
	const TCHAR PLUGIN_PREFIX_NAME[] = _T("{67850823-5AF3-4188-81F0-3FC00AF78281}+");
	Ostringstream os;
	os << dwThreadID;

	return PLUGIN_PREFIX_NAME + os.str();
}


CComQIPtr<IExplorerPlugin> MarshalService::FindBrwsPluginByThreadIDInTimeout(DWORD dwThreadID, DWORD nTimeout)
{
	CComQIPtr<IExplorerPlugin> spIePlugin;
	DWORD dwStartTime = ::GetTickCount();

	while (TRUE)
	{
		spIePlugin = FindBrwsPluginByThreadID(dwThreadID);
		if (spIePlugin != NULL)
		{
			// The plugin was found.
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

	return spIePlugin;
}


CComQIPtr<IExplorerPlugin> MarshalService::FindBrwsPluginByThreadID(DWORD dwThreadID)
{
	_ASSERTE(dwThreadID != 0);

	String sPluginMarshalName = GetMarshalWndName(dwThreadID);
	HWND   hPluginCommWnd     = ::FindWindow(HIDDEN_COMMUNICATION_WND_CLASS_NAME, sPluginMarshalName.c_str());

	if (NULL == hPluginCommWnd)
	{
		traceLog << "Can NOT find communication window in MarshalService::FindBrwsPluginByThreadID\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	CComQIPtr<IAccessible> spAccPlugin;
	HRESULT hRes = AccessibleObjectFromWindow(hPluginCommWnd, (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)&spAccPlugin);

	if (FAILED(hRes) || (spAccPlugin == NULL))
	{
		traceLog << "Can NOT get IAccessible object from wnd in MarshalService::FindBrwsPluginByThreadID\n";
		return CComQIPtr<IExplorerPlugin>();
	}

	CComQIPtr<IServiceProvider> spServProv = spAccPlugin;
	CComQIPtr<IExplorerPlugin> spPlugin;

	hRes = spServProv->QueryService(IID_IExplorerPlugin, IID_IExplorerPlugin, (void**)&spPlugin);
	return spPlugin;
}


void MarshalService::GetAllBrwsPluginThreads(std::list<DWORD>& pluginThreadsList)
{
	pluginThreadsList.clear();
	::EnumWindows(FindMarshallWndCallback, (LPARAM)(&pluginThreadsList));
}


BOOL CALLBACK MarshalService::FindMarshallWndCallback(HWND hWnd, LPARAM lParam)
{
	std::list<DWORD>* pList = reinterpret_cast<std::list<DWORD>*>(lParam);
	ATLASSERT(pList != NULL);

	if (::IsWindow(hWnd))
	{
		// Get the class name of the window.
		String sClassName = Common::GetWndClass(hWnd);
		if (HIDDEN_COMMUNICATION_WND_CLASS_NAME == sClassName)
		{
			DWORD dwThreadID = ::GetWindowThreadProcessId(hWnd, NULL);
			pList->push_back(dwThreadID);
		}
	}

	return TRUE;
}


void MarshalService::GetAllTabsPluginThreads(HWND hIEFrame, std::list<DWORD>& pluginThreadsList)
{
	if (GetWndClass(hIEFrame) != _T("IEFrame"))
	{
		return;
	}

	FIND_TABS_INFO findTabsInfo(hIEFrame, &pluginThreadsList);

	pluginThreadsList.clear();
	::EnumWindows(FindTabMarshallWndCallback, (LPARAM)(&findTabsInfo));
}


BOOL CALLBACK MarshalService::FindTabMarshallWndCallback(HWND hWnd, LPARAM lParam)
{
	FIND_TABS_INFO* pInfo = reinterpret_cast<FIND_TABS_INFO*>(lParam);
	ATLASSERT((pInfo != NULL) && (pInfo->m_pTabThreadList != NULL));
	ATLASSERT(GetWndClass(pInfo->m_hIEFrame) == _T("IEFrame"));

	if (::IsWindow(hWnd))
	{
		// Get the class name of the window.
		String sClassName = Common::GetWndClass(hWnd);
		if (HIDDEN_COMMUNICATION_WND_CLASS_NAME == sClassName)
		{
			// Communication windows responds to WM_APP_GET_IE_FRAME_MSG with their IEFrame handle.
			LRESULT lRes = ::SendMessage(hWnd, WM_APP_GET_IE_FRAME_MSG, 0, 0);
			if (lRes == HandleToLong(pInfo->m_hIEFrame))
			{
				DWORD dwThreadID = ::GetWindowThreadProcessId(hWnd, NULL);
				pInfo->m_pTabThreadList->push_back(dwThreadID);
			}
		}
	}

	return TRUE;
}


BOOL MarshalService::FindPluginThreadNotInList(HWND hIEFrame, std::list<DWORD>& pluginThreadsList, DWORD& dwNewThreadID, DWORD nTimeout)
{
	if (GetWndClass(hIEFrame) != _T("IEFrame"))
	{
		return FALSE;
	}

	DWORD dwStartTime = ::GetTickCount();

	while (TRUE)
	{
		FIND_TABS_INFO findNewTabInfo(hIEFrame, &pluginThreadsList);
		::EnumWindows(FindNewTabMarshallWndCallback, (LPARAM)(&findNewTabInfo));

		dwNewThreadID = findNewTabInfo.m_dwThreadResult;
		if (dwNewThreadID != 0)
		{
			// Thread ID found!
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


BOOL CALLBACK MarshalService::FindNewTabMarshallWndCallback(HWND hWnd, LPARAM lParam)
{
	FIND_TABS_INFO* pInfo = reinterpret_cast<FIND_TABS_INFO*>(lParam);
	ATLASSERT((pInfo != NULL) && (pInfo->m_pTabThreadList != NULL));
	ATLASSERT(GetWndClass(pInfo->m_hIEFrame) == _T("IEFrame"));

	if (::IsWindow(hWnd))
	{
		// Get the class name of the window.
		String sClassName = Common::GetWndClass(hWnd);
		if (HIDDEN_COMMUNICATION_WND_CLASS_NAME == sClassName)
		{
			// Communication windows responds to WM_APP_GET_IE_FRAME_MSG with their IEFrame handle.
			LRESULT lRes = ::SendMessage(hWnd, WM_APP_GET_IE_FRAME_MSG, 0, 0);
			if (lRes == HandleToLong(pInfo->m_hIEFrame))
			{
				DWORD dwThreadID = ::GetWindowThreadProcessId(hWnd, NULL);

				std::list<DWORD>::const_iterator it = std::find(pInfo->m_pTabThreadList->begin(), pInfo->m_pTabThreadList->end(), dwThreadID);
				if (it == pInfo->m_pTabThreadList->end())
				{
					pInfo->m_dwThreadResult = dwThreadID;
					return FALSE;
				}
			}
		}
	}

	return TRUE;
}
