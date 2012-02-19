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

#pragma once
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "Common.h"
using namespace Common;


namespace NewMarshalService
{
	const TCHAR HIDDEN_COMMUNICATION_WND_CLASS_NAME[] = _T("{E7DF328E-5559-4c49-BB60-58114B1E467A}");

	CComQIPtr<IExplorerPlugin> FindBrwsPluginByIeWnd     (HWND hIeWnd, DWORD nTimeout = 0, BOOL bStopBrowser = FALSE, CComQIPtr<IWebBrowser2> spBrowser = CComQIPtr<IWebBrowser2>());
	BOOL                       GetMarshalWndName         (HWND hIeServer, Common::String& sRetName);
	void                       GetTabsIEServer           (HWND hIEFrame, std::list<HWND>& ieServerList);
	BOOL                       FindIeWndNotInList        (HWND hIEFrame, std::list<HWND>& ieServerList, HWND& hNewThreadWnd, DWORD nTimeout = Common::INTERNAL_GLOBAL_TIMEOUT);
	BOOL                       FindChildShDocViewCallback(HWND hWnd, void*);

	void GetIEServerList(std::list<HWND>& ieServerList, const String& sAppName, BOOL bEqual);
}
