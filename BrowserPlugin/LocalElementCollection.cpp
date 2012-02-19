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

#include "stdafx.h"
#include "DebugServices.h"
#include "LocalElementCollection.h"
#undef min


const int NUMBER_OF_ELEMENTS_IN_PAGE = 500;

// CLocalElementCollection

void CLocalElementCollection::InitCollection(const list<CAdapt<CComQIPtr<IHTMLElement> > >& elementList)
{
	ATLASSERT(m_elementCollection.size() == 0);
	size_t nNumberOfElements = elementList.size();

	if (0 == nNumberOfElements)
	{
		return;
	}

	m_elementCollection.reserve(nNumberOfElements);
	for (list<CAdapt<CComQIPtr<IHTMLElement> > >::const_iterator it = elementList.begin();
		 it != elementList.end(); ++it)
	{
		m_elementCollection.push_back(*it);
	}

	ATLASSERT(NUMBER_OF_ELEMENTS_IN_PAGE > 0);
	m_nNumberOfElements = static_cast<long>(m_elementCollection.size());
	m_nNumberOfPages    = (m_nNumberOfElements / NUMBER_OF_ELEMENTS_IN_PAGE) + 1;
}


STDMETHODIMP CLocalElementCollection::GetCollectionInfo(LONG* pNumberOfElements, LONG* pNumberOfPages, LONG* pPageSize)
{
	if (NULL == pNumberOfElements)
	{
		traceLog << "pNumberOfElements is NULL in CLocalElementCollection::GetCollectionInfo\n";
		return E_INVALIDARG;
	}

	if (NULL == pNumberOfPages)
	{
		traceLog << "pNumberOfPages is NULL in CLocalElementCollection::GetCollectionInfo\n";
		return E_INVALIDARG;
	}

	if (NULL == pPageSize)
	{
		traceLog << "pPageSize is NULL in CLocalElementCollection::GetCollectionInfo\n";
		return E_INVALIDARG;
	}

	ATLASSERT(m_nNumberOfElements >= 0);
	ATLASSERT(m_nNumberOfPages    >= 0);

	*pNumberOfElements = m_nNumberOfElements;
	*pNumberOfPages    = m_nNumberOfPages;
	*pPageSize         = NUMBER_OF_ELEMENTS_IN_PAGE;

	return S_OK;
}


STDMETHODIMP CLocalElementCollection::GetPage(LONG nPageNumber, SAFEARRAY** pPages)
{
	if ((m_nNumberOfPages - 1 < nPageNumber) || (nPageNumber < 0))
	{
		traceLog << "Invalid page number " << nPageNumber << "in CLocalElementCollection::GetPage\n";
		traceLog << "Max page number is " << m_nNumberOfPages << "in CLocalElementCollection::GetPage\n";
		return E_INVALIDARG;
	}

	if (NULL == pPages)
	{
		traceLog << "pPages is null in CLocalElementCollection::GetPage\n";
		return E_INVALIDARG;
	}

	long nFirstIndex       = nPageNumber * NUMBER_OF_ELEMENTS_IN_PAGE;
	long nLastIndex        = std::min(nFirstIndex + NUMBER_OF_ELEMENTS_IN_PAGE - 1, m_nNumberOfElements - 1);
	long nNumberOfElements = nLastIndex - nFirstIndex + 1;

	// Create a safe array.
	SAFEARRAY*     psa;
	SAFEARRAYBOUND rgsabound[1];

	rgsabound[0].lLbound   = 0;
	rgsabound[0].cElements = nNumberOfElements;

	psa = ::SafeArrayCreate(VT_DISPATCH, 1, rgsabound);

	if (NULL == psa)
	{
		traceLog << "SafeArrayCreate failed in CLocalElementCollection::GetPage\n";
		return E_OUTOFMEMORY;
	}

	LONG j = 0;
	for (int i = nFirstIndex; i <= nLastIndex; ++i)
	{
		CComQIPtr<IDispatch> spDisp = m_elementCollection[i];
		HRESULT hr = ::SafeArrayPutElement(psa, &j, spDisp.p); // I assume it calls AddRef.

		j++;
	}

	*pPages = psa;
	return S_OK;
}
