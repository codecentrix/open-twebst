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


// CLocalElementCollection

class ATL_NO_VTABLE CLocalElementCollection : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CLocalElementCollection>,
	public IDispatchImpl<ILocalElementCollection, &IID_ILocalElementCollection, &LIBID_OpenTwebstPluginLib>
{
public:
	CLocalElementCollection()
	{
		m_nNumberOfPages    = 0;
		m_nNumberOfElements = 0;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_LOCALELEMENTCOLLECTION)

DECLARE_NOT_AGGREGATABLE(CLocalElementCollection)

BEGIN_COM_MAP(CLocalElementCollection)
	COM_INTERFACE_ENTRY(ILocalElementCollection)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease() 
	{
	}

	STDMETHOD(GetCollectionInfo)(LONG* pNumberOfElements, LONG* pNumberOfPages, LONG* pPageSize);
	STDMETHOD(GetPage)          (LONG nPageNumber, SAFEARRAY** pPages);

public:
	void InitCollection(const list<CAdapt<CComQIPtr<IHTMLElement> > >& elementList);


private:
	vector<CComQIPtr<IHTMLElement> > m_elementCollection;
	long                             m_nNumberOfPages;
	long                             m_nNumberOfElements;
};
