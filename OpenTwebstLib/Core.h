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
#include "resource.h"       // main symbols
#include "Common.h"
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "CatCLSLib.h"


// CCore

class ATL_NO_VTABLE CCore : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCore, &CLSID_Core>,
	public ISupportErrorInfo,
	public IDispatchImpl<ICore, &IID_ICore, &LIBID_OpenTwebstLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IConnectionPointContainerImpl<CCore>,
	public IConnectionPointImpl<CCore, &DIID__ICoreEvents>
{
public:
	CCore();
	void    SetLastErrorCode(LONG nErr);
	BOOL    IsCancelPending();
	HRESULT Fire_CancelRequest();
	BOOL    GetAutoClosePopups();


	DECLARE_REGISTRY_RESOURCEID(IDR_CORE)

	BEGIN_COM_MAP(CCore)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(ICore)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CCore)
		CONNECTION_POINT_ENTRY(DIID__ICoreEvents)
	END_CONNECTION_POINT_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease() 
	{
	}

public:
	// ICore
	STDMETHOD(StartBrowser)                     (BSTR bstrUrl, IBrowser** ppNewBrowser);
	STDMETHOD(FindBrowser)                      (BSTR bstrCond, IBrowser** ppBrowserFound);
	STDMETHOD(get_useHardwareInputEvents)       (VARIANT_BOOL* pVal);
	STDMETHOD(put_useHardwareInputEvents)       (VARIANT_BOOL newVal);
	STDMETHOD(get_searchTimeout)                (LONG* pVal);
	STDMETHOD(put_searchTimeout)                (LONG newVal);
	STDMETHOD(get_loadTimeout)                  (LONG* pVal);
	STDMETHOD(put_loadTimeout)                  (LONG newVal);
	STDMETHOD(get_lastError)                    (LONG* pVal);
	STDMETHOD(get_loadTimeoutIsError)           (VARIANT_BOOL* pVal);
	STDMETHOD(put_loadTimeoutIsError)           (VARIANT_BOOL newVal);
	STDMETHOD(AttachToHWND)                     (LONG nWnd, IBrowser** ppBrowser);
	STDMETHOD(AttachToNativeFrame)              (IHTMLWindow2* pHtmlWindow, IFrame** ppFrame);
	STDMETHOD(AttachToNativeElement)            (IHTMLElement* pHtmlElement, IElement** ppElement);
	STDMETHOD(AttachToNativeBrowser)            (IWebBrowser2* pWebBrowser,  IBrowser** ppBrowser);
	STDMETHOD(Reset)                            ();
	STDMETHOD(FindElementFromPoint)             (LONG x, LONG y, IElement** ppElement);
	STDMETHOD(get_INDEX_OUT_OF_BOUND_ERROR)     (LONG* pVal);
	STDMETHOD(get_BROWSER_CONNECTION_LOST_ERROR)(LONG* pVal);
	STDMETHOD(get_INVALID_OPERATION_ERROR)      (LONG* pVal);
	STDMETHOD(get_NOT_FOUND_ERROR)              (LONG* pVal);
	STDMETHOD(get_OK_ERROR)                     (LONG* pVal);
	STDMETHOD(get_FAIL_ERROR)                   (LONG* pVal);
	STDMETHOD(get_INVALID_ARG_ERROR)            (LONG* pVal);
	STDMETHOD(get_LOAD_TIMEOUT_ERROR)           (LONG* pVal);
	STDMETHOD(get_IEVersion)                    (BSTR* pBstrVersion);
	STDMETHOD(get_foregroundBrowser)            (IBrowser** pFgBrowser);
	STDMETHOD(get_productVersion)               (BSTR* pBstrVersion);
	STDMETHOD(get_productName)                  (BSTR* pBstrName);
	STDMETHOD(GetClipboardText)                 (BSTR* pBstrClipboardText);
	STDMETHOD(SetClipboardText)                 (BSTR bstrClipboardText);
	STDMETHOD(get_closeBrowserPopups)           (VARIANT_BOOL* pAutoClosePopups);
	STDMETHOD(put_closeBrowserPopups)           (VARIANT_BOOL  vbAutoClosePopups);
	STDMETHOD(get_asyncHtmlEvents)              (VARIANT_BOOL* pVal);
	STDMETHOD(put_asyncHtmlEvents)              (VARIANT_BOOL newVal);


private:
	static BOOL                FindVisibleChildShDocViewCallback(HWND hWnd, void* pObj);
	static BOOL                FindVisibleChildWndCallback(HWND hWnd, void* pObj);
	static BOOL CALLBACK       FindWndCallback(HWND hWnd, LPARAM lParam);
	static BOOL CALLBACK       FindWndIE8OnVistaCallback(HWND hWnd, LPARAM lParam);
	static BOOL CALLBACK       FindTabWndCallback(HWND hWnd, LPARAM lParam);
	static BOOL CALLBACK       FindTabWndIE8OnVistaCallback(HWND hWnd, LPARAM lParam);
	static void                FindTabWndIE8OnVista(HWND hIEFrameWnd, LPARAM lParam);

	DWORD                      StartInternetExplorerProcess(BOOL bStartHidden, BSTR bstrUrl);
	DWORD                      StartInternetExplorerProcessOnVista(LPCWSTR pszUrl, BOOL bStartHidden);
	HWND                       FindIeWnd(DWORD dwIEProcID);
	HWND                       FindIeWndOnVistaIE8(DWORD dwIEProcID);
	HWND                       FindTabWnd(HWND hIEFrameWnd);
	HRESULT                    StartBrowser(IBrowser** ppNewBrowser, DWORD dwHelpID, BOOL bStartHidden, BSTR bstrUrl);
	CComQIPtr<IBrowser>        FindBrowser(SAFEARRAY* pVarArgs, const String& sAppName, BOOL bUseEqOp);
	CComQIPtr<IExplorerPlugin> FindBrwsPluginFromHtmlWindow(IHTMLWindow2* pHtmlWindow);
	CComQIPtr<IExplorerPlugin> FindBrwsPluginFromWebBrowser(IWebBrowser2* pWebBrowser);
	CComQIPtr<IExplorerPlugin> FindBrwsPluginFromHtmlElement(IHTMLElement* pHtmlElement);
	BOOL                       FindBrowserList(SAFEARRAY* pVarArgs, std::list<CAdapt<CComQIPtr<IExplorerPlugin> > >* pBrowserList, const String& sAppName, BOOL bUseEqOp);
	HRESULT                    SetComErrorMessage(UINT nId, DWORD dwHelpID);
	HRESULT                    SetComErrorMessage(UINT nIdPattern, LPCTSTR szFunctionName, DWORD dwHelpID);
	BOOL                       IsValidDescriptorList(const std::list<DescriptorToken>& tokens);
	BOOL                       GetAppFilterData(const std::list<DescriptorToken>&tokenList, String& sOutAppName, BOOL& bOutUseEqOp);
	HRESULT                    ClearClipboard();
	BOOL                       SetTextInClipboard(LPCWSTR szText);
	BOOL                       SetTextInClipboard(LPCSTR szText);
	void                       FindAllIEFramesWnd(std::set<HWND>& allIeFrames);
	void                       ProcessMessagesOrSleep();
	BOOL                       AttachToIEFrame(HWND hIEFrameWnd, IBrowser** ppBrowser);
	BOOL                       AttachToTab(HWND hIETabWnd, IBrowser** ppBrowser);
	BOOL                       AttachToIEServer(HWND hIEServerWnd, IBrowser** ppBrowser);
	BOOL                       AttachToTridentDlg(HWND hTargetWnd, IBrowser** ppBrowser);
	HWND                       FindIEFrameExcludeSet(std::set<HWND>& ieFramesToExclude);
	static BOOL CALLBACK       FindAllIEFramesWndCallback(HWND hWnd, LPARAM lParam);
	static BOOL CALLBACK       FindIEFrameExcludeSetCallback(HWND hWnd, LPARAM lParam);


private:
	VARIANT_BOOL m_bAsyncEvents;
	BOOL         m_bCanceled;
	LONG         m_nSearchTimeout;
	LONG         m_nLoadTimeout;
	LONG         m_nLastError;
	ULONG        m_nSettingsFlag;
	VARIANT_BOOL m_vbLoadTimeoutIsErr;
	VARIANT_BOOL m_vbUseHardwareEvents;
	VARIANT_BOOL m_vbAutoClosePopups;
};

OBJECT_ENTRY_AUTO(__uuidof(Core), CCore)
