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


// CFrame

class ATL_NO_VTABLE CFrame : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFrame>,
	public ISupportErrorInfo,
	public IDispatchImpl<IFrame, &IID_IFrame, &LIBID_OpenTwebstLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public BaseLibObject
{
public:
	CFrame() : BaseLibObject(GetObjectCLSID())
	{
	}

	void SetHtmlWindow(IHTMLWindow2* pHtmlWindow);


DECLARE_REGISTRY_RESOURCEID(IDR_FRAME)


BEGIN_COM_MAP(CFrame)
	COM_INTERFACE_ENTRY(IFrame)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	// IFrame
	STDMETHOD(get_parentBrowser)   (IBrowser**       pVal);
	STDMETHOD(get_nativeFrame)     (IHTMLWindow2**   pVal);
	STDMETHOD(get_document)        (IHTMLDocument2** pVal);
	STDMETHOD(get_core)            (ICore**          pVal);
	STDMETHOD(FindElement)         (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(FindChildElement)    (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(FindFrame)           (BSTR bstrCond, IFrame** ppFrame);
	STDMETHOD(FindChildFrame)      (BSTR bstrCond, IFrame** ppFrame);
	STDMETHOD(FindAllElements)     (BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList);
	STDMETHOD(FindChildrenElements)(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList);
	STDMETHOD(get_parentFrame)     (IFrame** ppParentFrame);

	/////////////////////////////////////////////////////////////////////////////////////////
	// Methods and properties not yet implemented.
	STDMETHOD(get_frameElement)    (IElement** ppFrameElement);
	STDMETHOD(get_title)           (BSTR* pVal);
	STDMETHOD(get_url)             (BSTR* pVal);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

protected:
	virtual BOOL IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID);

private:
	CComQIPtr<IHTMLWindow2> m_spHtmlWindow;
};
