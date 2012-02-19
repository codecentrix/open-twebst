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

// BrowserPlugin.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "DebugServices.h"
#include "BrowserPlugin.h"
#include "dlldatax.h"


class CBrowserPluginModule : public CAtlDllModuleT< CBrowserPluginModule >
{
public :
	DECLARE_LIBID(LIBID_OpenTwebstPluginLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_BROWSERPLUGIN, "{E669D391-C510-4c2d-9E12-FD7B436D1396}")
};

CBrowserPluginModule _AtlModule;


// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif

	hInstance;
	switch (dwReason)
	{
		case DLL_PROCESS_ATTACH :
		{
			// Get the name of the executable where the BHO is injected.
			TCHAR szCurrentExeName[MAX_PATH + 1];

			::GetModuleFileName(NULL, szCurrentExeName, MAX_PATH);
			szCurrentExeName[MAX_PATH] = _T('\0');

			// Get the short exe name.
			TCHAR* szShortName = _tcsrchr(szCurrentExeName, _T('\\'));
			szShortName = (szShortName ? szShortName + 1 : szCurrentExeName);
			traceLog << "BHO loading in: '" << szShortName << "'\n";

			if (!_tcsicmp( _T("explorer.exe"), szShortName))
			{
				// Don't allow "Windows Explorer" to load the BHO.
				traceLog << "Don't load the BHO in explorer.exe\n";
				return FALSE;
			}

			break;
		}
	};

    return _AtlModule.DllMain(dwReason, lpReserved); 
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
    HRESULT hr = PrxDllCanUnloadNow();
    if (hr != S_OK)
        return hr;
#endif

    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
        return S_OK;
#endif

    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
#ifdef _MERGE_PROXYSTUB
    if (FAILED(hr))
        return hr;
    hr = PrxDllRegisterServer();
#endif

	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
#ifdef _MERGE_PROXYSTUB
    if (FAILED(hr))
        return hr;
    hr = PrxDllRegisterServer();
    if (FAILED(hr))
        return hr;
    hr = PrxDllUnregisterServer();
#endif

	return hr;
}


STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
    HRESULT hr = E_FAIL;
    static const wchar_t szUserSwitch[] = _T("user");

    if (pszCmdLine != NULL)
    {
    	if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
    	{
    		AtlSetPerUserRegistration(true);
    	}
    }

    if (bInstall)
    {	
    	hr = DllRegisterServer();
    	if (FAILED(hr))
    	{	
    		DllUnregisterServer();
    	}
    }
    else
    {
    	hr = DllUnregisterServer();
    }

    return hr;
}


STDAPI DllRegisterPerUser()
{
	return DllInstall(TRUE, _T("user"));
}


STDAPI DllUnRegisterPerUser()
{
	return DllInstall(FALSE, _T("user"));
}
