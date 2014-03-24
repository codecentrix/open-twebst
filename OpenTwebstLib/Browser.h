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
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "CatCLSLib.h"
#include "BaseLibObject.h"


// CBrowser

class ATL_NO_VTABLE CBrowser : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CBrowser>,
	public ISupportErrorInfo,
	public IDispatchImpl<IBrowser, &IID_IBrowser, &LIBID_OpenTwebstLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public BaseLibObject
{
public:
	CBrowser() : BaseLibObject(GetObjectCLSID())
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_BROWSER)


BEGIN_COM_MAP(CBrowser)
	COM_INTERFACE_ENTRY(IBrowser)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

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
	// IBrowser methods.
	STDMETHOD(Navigate)              (BSTR bstrUrl);
	STDMETHOD(WaitToLoad)            (BSTR bstrCond, VARIANT_BOOL* pVal);
	STDMETHOD(FindElement)           (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(FindFrame)             (BSTR bstrCond, IFrame** ppFrame);
	STDMETHOD(FindAllElements)       (BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList);
	STDMETHOD(Close)                 (void);
	STDMETHOD(get_nativeBrowser)     (IWebBrowser2** pVal);
	STDMETHOD(get_isLoading)         (VARIANT_BOOL* pVal);
	STDMETHOD(get_title)             (BSTR* pVal);
	STDMETHOD(get_url)               (BSTR* pVal);
	STDMETHOD(get_app)               (BSTR* pVal);
	STDMETHOD(get_core)              (ICore** pVal);
	STDMETHOD(get_topFrame)          (IFrame** ppTopFrame);
	STDMETHOD(get_navigationError)   (LONG* pVal);
	STDMETHOD(FindModelessHtmlDialog)(BSTR bstrCond, IFrame** ppFrame);
	STDMETHOD(FindModalHtmlDialog)   (IFrame** ppFrame);
	STDMETHOD(ClosePopup)            (BSTR bstrPopupText, VARIANT vButton, BSTR* pPopupText);
	STDMETHOD(ClosePrompt)           (BSTR bstrPromptText, BSTR bstrValue, VARIANT vButton, BSTR* pPopupText);
	STDMETHOD(GetAttr)               (BSTR bstrAttrName, VARIANT* pVal);
	STDMETHOD(SetAttr)               (BSTR bstrAttrName, VARIANT newVal);

	/////////////////////////////////////////////////////////////////////////////////////////
	// Methods and properties not yet implemented.
	STDMETHOD(SaveSnapshot)               (BSTR bstrImgFileName);
	STDMETHOD(FindElementByXPath)         (BSTR bstrXPath, IElement** ppElement);

private:
	HRESULT ClosePopupOrPrompt   (BSTR bstrPopupText, VARIANT vButton, BSTR bstrValue, BSTR* pPopupText);
	HRESULT GetWebBrowser        (IWebBrowser2** ppWebBrowser);
	BOOL    IsValidDescriptorList(const std::list<DescriptorToken>& tokens);

	HRESULT GetAppName(VARIANT* pAppName);
	HRESULT GetPID    (VARIANT* pPid);
};
