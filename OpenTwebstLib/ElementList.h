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
#include "Element.h"
#include "CatCLSLib.h"


// CElementList

class ATL_NO_VTABLE CElementList : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CElementList>,
	public ISupportErrorInfo,
	public IDispatchImpl<IElementList, &IID_IElementList, &LIBID_OpenTwebstLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public BaseLibObject
{
public:
	CElementList() : BaseLibObject(GetObjectCLSID())
	{
		m_nNumberOfElements = 0;
		m_nNumberOfPages    = 0;
		m_nMaxPageSize      = 0;
	}

	~CElementList();
	HRESULT InitElementList(ILocalElementCollection* pElementCollection);


DECLARE_REGISTRY_RESOURCEID(IDR_ELEMENTLIST)


BEGIN_COM_MAP(CElementList)
	COM_INTERFACE_ENTRY(IElementList)
	COM_INTERFACE_ENTRY2(IDispatch, IElementList)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);
	STDMETHOD(get_length)                (LONG* pVal);
	STDMETHOD(get_item)                  (LONG nIndex, IElement** pVal);
	STDMETHOD(get_core)                  (ICore** pVal);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease() 
	{
	}


private:
	CComQIPtr<IElement> GetElement(LONG nIndex);

private:
	struct ElementPage
	{
		ElementPage()
		{
			m_nPageSize       = 0;
			m_psaPageElements = NULL;
		}

		LONG                              m_nPageSize;
		SAFEARRAY*                        m_psaPageElements;
		std::vector<CComQIPtr<IElement> > m_elements;
	};

private:
	CComQIPtr<ILocalElementCollection> m_spElementCollection;
	LONG                               m_nNumberOfElements;
	LONG                               m_nNumberOfPages;
	LONG                               m_nMaxPageSize;
	std::vector<ElementPage>           m_pages;
};
