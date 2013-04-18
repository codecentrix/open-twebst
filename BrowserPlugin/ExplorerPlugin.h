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
#include "BrowserPlugin.h"
#include "MarshalService.h"
#include "AsyncAction.h"
using namespace std;


#define WM_APP_ASYNC_ACTION (WM_APP + 3)


// CExplorerPlugin
class ATL_NO_VTABLE CExplorerPlugin : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExplorerPlugin, &CLSID_OpenTwebstBHO>,
	public IObjectWithSiteImpl<CExplorerPlugin>,
	public IDispatchImpl<IAccessible, &IID_IAccessible, &LIBID_Accessibility>,
	public IDispatchImpl<IExplorerPlugin, &IID_IExplorerPlugin, &LIBID_OpenTwebstPluginLib>,
	public CWindowImpl<CExplorerPlugin>,
	public IServiceProviderImpl<CExplorerPlugin>
{
friend class SelectAsyncAction;

public:
	CExplorerPlugin()
	{
		m_nLastNavigationErr = 0;
		m_dwEventsCookie     = 0;
		m_nActiveDownloads   = 0;
		m_bIsForceLoaded     = FALSE;
		m_hIeWnd             = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERPLUGIN)

DECLARE_NOT_AGGREGATABLE(CExplorerPlugin)

BEGIN_COM_MAP(CExplorerPlugin)
	COM_INTERFACE_ENTRY(IExplorerPlugin)
	COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY2(IDispatch, IExplorerPlugin)
	COM_INTERFACE_ENTRY(IAccessible)
	COM_INTERFACE_ENTRY(IServiceProvider)
END_COM_MAP()


BEGIN_SERVICE_MAP(CExplorerPlugin)
	SERVICE_ENTRY(IID_IExplorerPlugin)
END_SERVICE_MAP()

	DECLARE_WND_CLASS(NewMarshalService::HIDDEN_COMMUNICATION_WND_CLASS_NAME);
	DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_MSG_MAP(CNetIeBHO)
	MESSAGE_HANDLER(WM_GETOBJECT,                            OnAccGetObject)
	MESSAGE_HANDLER(WM_APP_ASYNC_ACTION,                     OnAsyncAction)
	//MESSAGE_HANDLER(MarshalService::WM_APP_GET_IE_FRAME_MSG, OnGetIEFrame)

END_MSG_MAP()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease() 
	{
	}


//Message handling
private:
	LRESULT OnAccGetObject(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAsyncAction (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//LRESULT OnGetIEFrame  (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

// IAccessible.
public:
	STDMETHOD(get_accParent)          (IDispatch **ppdispParent);
	STDMETHOD(get_accChildCount)      (long *pcountChildren);
	STDMETHOD(get_accChild)           (VARIANT varChild, IDispatch **ppdispChild);
	STDMETHOD(get_accName)            (VARIANT varChild, BSTR *pszName);
	STDMETHOD(get_accValue)           (VARIANT varChild, BSTR *pszValue);
	STDMETHOD(get_accDescription)     (VARIANT varChild, BSTR *pszDescription);
	STDMETHOD(get_accRole)            (VARIANT varChild, VARIANT *pvarRole);
	STDMETHOD(get_accState)           (VARIANT varChild, VARIANT *pvarState);
	STDMETHOD(get_accHelp)            (VARIANT varChild, BSTR *pszHelp);
	STDMETHOD(get_accHelpTopic)       (BSTR *pszHelpFile, VARIANT varChild, long *pidTopic);
	STDMETHOD(get_accKeyboardShortcut)(VARIANT varChild, BSTR *pszKeyboardShortcut);
	STDMETHOD(get_accFocus)           (VARIANT *pvarChild);
	STDMETHOD(get_accSelection)       (VARIANT *pvarChildren);
	STDMETHOD(get_accDefaultAction)   (VARIANT varChild, BSTR *pszDefaultAction);
	STDMETHOD(accSelect)              (long flagsSelect, VARIANT varChild);
	STDMETHOD(accLocation)            (long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild);
	STDMETHOD(accNavigate)            (long navDir, VARIANT varStart, VARIANT *pvarEndUpAt);
	STDMETHOD(accHitTest)             (long xLeft, long yTop, VARIANT *pvarChild);
	STDMETHOD(accDoDefaultAction)     (VARIANT varChild);
	STDMETHOD(put_accName)            (VARIANT varChild, BSTR szName);
	STDMETHOD(put_accValue)           (VARIANT varChild, BSTR szValue);


public:
	// IExplorerPlugin interface.
	STDMETHOD(GetDocumentFromWindow)  (IHTMLWindow2* pWindow, IHTMLDocument2** ppDocument);
	STDMETHOD(GetElementScreenRect)   (IHTMLElement* pElement, long* pLeft, long* pTop, long* pRight, long* pBottom);
	STDMETHOD(GetLastNavigationErr)   (LONG* pErrCode);
	STDMETHOD(GetIEServerWnd)         (IHTMLElement* pElement, LONG* pWnd);
	STDMETHOD(GetElementText)         (IHTMLElement* pElement, BSTR* pBstrText);
	STDMETHOD(GetNativeBrowser)       (IWebBrowser2** ppWebBrowser);
	STDMETHOD(CheckBrowserDescriptor) (LONG nSearchFlags, SAFEARRAY * psa);
	STDMETHOD(SetSite)                (IUnknown* pSite);
	STDMETHOD(IsLoading)              (VARIANT_BOOL* pLoading);
	STDMETHOD(GetTopFrame)            (IHTMLWindow2** pTopFrame);
	STDMETHOD(FindInContainer)        (IUnknown* pContainer, LONG nSearchFlags, BSTR bstrTagName, SAFEARRAY* psa, IUnknown** ppResult);
	STDMETHOD(FindFramesInContainer)  (IUnknown* pContainer, LONG nSearchFlags, SAFEARRAY* psa, LONG* pSize, IHTMLWindow2*** pppFrames);
	STDMETHOD(GetScreenClickPoint)    (IHTMLElement* pElement, LONG nGetInputFileButton, LONG* pX, LONG* pY);
	STDMETHOD(SelectOptions)          (IHTMLElement* pElement, VARIANT vStart, VARIANT vEnd, LONG nFlags);
	STDMETHOD(ClearSelection)         (IHTMLElement* pElement);
	STDMETHOD(GetBrowserTitle)        (BSTR* pBstrTitle);
	STDMETHOD(GetAppName)             (BSTR* pBstrApp);
	STDMETHOD(GetBrowserThreadID)     (LONG* pThID);
	STDMETHOD(GetBrowserProcessID)    (LONG* pProcID);
	STDMETHOD(SetFocusOnElement)      (IHTMLElement* pTargetElement, VARIANT_BOOL vbAsync);
	STDMETHOD(SetFocusAwayFromElement)(IHTMLElement* pTargetElement, BOOL bGenerateOnChange, VARIANT_BOOL vbAsync);
	STDMETHOD(FireEventOnElement)     (IHTMLElement* pTargetElement, BSTR bstrEventName, LONG wCharToRise, VARIANT_BOOL vbAsync);
	STDMETHOD(ClickElementAt)         (IHTMLElement* pElement, LONG x, LONG y, LONG nRightClick);
	STDMETHOD(ClickElementAtAsync)    (IHTMLElement* pElement, LONG x, LONG y, LONG nRightClick);
	STDMETHOD(PostInputText)          (LONG nIEWnd, BSTR bstrText, BOOL bIsInputFile);
	STDMETHOD(FindSelectedOption)     (IHTMLElement* pElement, IHTMLElement** ppSelectedOption);
	STDMETHOD(FindAllSelectedOptions) (IHTMLElement* pElement, IUnknown** ppResult);
	STDMETHOD(SetForceLoaded)         (LONG nIeWnd);
	STDMETHOD(FindElementFromPoint)   (LONG x, LONG y, IHTMLElement** ppElem);

	// IDispatch interface.
	STDMETHOD(Invoke)(DISPID dispidMember,REFIID riid, LCID lcid, WORD nFlags,
					  DISPPARAMS * pdispparams, VARIANT * pvarResult,
					  EXCEPINFO * pexcepinfo, UINT * puArgErr);

private:
	HRESULT Init(IUnknown* pSite);
	HRESULT UnInit();
	void    OnQuit();

	BOOL GetBrowserTitle         (String* pTitle);
	BOOL GetBrowserUrl           (String* pUrl);
	void GetBrowserPid           (String* pPid);
	void GetBrowserTid           (String* pTid);
	BOOL GetBrowserAttributeValue(const String& sAttributeName, String* pAttributeValue);
	BOOL SelectComboItem         (CComQIPtr<IHTMLSelectElement> spSelectElem, DWORD dwItem);
	BOOL SelectSingleListItem    (CComQIPtr<IHTMLSelectElement> spSelectElem,  DWORD dwItem);
	BOOL SelectMultipleListItem  (CComQIPtr<IHTMLSelectElement> spSelectElem, const list<DWORD>& itemsToSelect);
	void ClearSelection          (CComQIPtr<IHTMLSelectElement> spSelectElem, BOOL bFireEvent);
	BOOL AddToSelection          (CComQIPtr<IHTMLSelectElement> spSelectElem, const list<DWORD>& itemsToSelect);
	void CleanAyncActionsQueue   ();
	int  GetZoomLevel            ();

	list<DWORD> FindItemsToSelect(VARIANT vItems, IHTMLElement* pElement, LONG nFlags);
	list<DWORD> FindItemsToSelect(VARIANT vStart, VARIANT vEnd, IHTMLElement* pElement, LONG nFlags);
	HRESULT     IsEmptyCollection(CComQIPtr<ILocalElementCollection> spColl, BOOL& bIsEmpty);

	HRESULT FindInfoForSelect(IHTMLElement* pElement, VARIANT vStart, VARIANT vEnd, LONG nFlags, list<DWORD>& outItemsToSelect, BOOL& bOutIsCombo, BOOL& bOutMultiple);
	HRESULT DoSelectOptions  (IHTMLElement* pElement, const list<DWORD>& itemsToSelect, BOOL bIsCombo, BOOL bMultiple, LONG nFlags);

	CComQIPtr<IHTMLDocument2> FindDocumentFromPoint(LONG x, LONG y);

private:
	DWORD               m_dwEventsCookie;
	LONG                m_nLastNavigationErr;
	LONG                m_nActiveDownloads;
	list<IAsyncAction*> m_asyncActionsQueue;
	BOOL                m_bIsForceLoaded;
	HWND                m_hIeWnd;
};

OBJECT_ENTRY_AUTO(__uuidof(OpenTwebstBHO), CExplorerPlugin)
