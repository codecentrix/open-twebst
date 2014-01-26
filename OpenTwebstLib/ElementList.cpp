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
#include "..\Common\CodeErrors.h"
#include "Core.h"
#include "HtmlHelpIDs.h"
#include "ElementList.h"
#include "MethodAndPropertyNames.h"

using namespace std;

// CElementList

STDMETHODIMP CElementList::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IElementList
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}


HRESULT CElementList::InitElementList(ILocalElementCollection* pElementCollection)
{
	ATLASSERT(pElementCollection    != NULL);
	ATLASSERT(m_spElementCollection == NULL);

	m_spElementCollection = pElementCollection;
	HRESULT hRes = m_spElementCollection->GetCollectionInfo(&m_nNumberOfElements, &m_nNumberOfPages, &m_nMaxPageSize);

	ATLASSERT(m_nNumberOfElements >= 0);
	ATLASSERT(m_nNumberOfPages    >= 0);
	ATLASSERT(m_nMaxPageSize      >  0);

	m_pages.resize(m_nNumberOfPages);
	return hRes;
}


STDMETHODIMP CElementList::get_length(LONG* pVal)
{
	FIRE_CANCEL_REQUEST();

	ATLASSERT(m_spElementCollection != NULL);
	ATLASSERT(m_spCore              != NULL);

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CElementList::get_length\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, LENGTH_PROPERTY, IDH_ELEMENT_LIST_COUNT);	// Set the error message.
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	*pVal = m_nNumberOfElements;
	return HRES_OK;
}


STDMETHODIMP CElementList::get_item(LONG nIndex, IElement** pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (m_spElementCollection == NULL)
	{
		traceLog << "m_spElementCollection is NULL in CElementList::get_item\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ITEM_PROPERTY, IDH_ELEMENT_LIST_ITEM);	// Set the error message.
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CElementList::get_item\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ITEM_PROPERTY, IDH_ELEMENT_LIST_ITEM);	// Set the error message.
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in CElementList::get_item\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ITEM_PROPERTY, IDH_ELEMENT_LIST_ITEM);	// Set the error message.
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	if (nIndex > m_nNumberOfElements - 1)
	{
		traceLog << "index out of bounds in CElementList::get_item\n";
		SetComErrorMessage(IDS_ERR_INDEX_OUT_OF_BOUNDS, IDH_ELEMENT_LIST_ITEM);	// Set the error message.
		SetLastErrorCode(ERR_INDEX_OUT_OF_BOUND);
		return HRES_INDEX_OUT_OF_BOUND_ERR;
	}

	CComQIPtr<IElement> spElement = GetElement(nIndex);
	if (spElement != NULL)
	{
		*pVal = spElement.Detach();
		return HRES_OK;
	}
	else
	{
		traceLog << "GetElement failed in CElementList::get_item\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ITEM_PROPERTY, IDH_ELEMENT_LIST_ITEM);	// Set the error message.
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
}


CElementList::~CElementList()
{
	for (vector<ElementPage>::iterator it = m_pages.begin();
		 it != m_pages.end(); ++it)
	{
		if (it->m_psaPageElements != NULL)
		{
			// Release is called on each object in the array.
			::SafeArrayDestroy(it->m_psaPageElements);
		}

		it->m_psaPageElements = NULL;
		it->m_nPageSize       = 0;
	}
}


CComQIPtr<IElement> CElementList::GetElement(LONG nIndex)
{
	ATLASSERT(nIndex < m_nNumberOfElements);
	ATLASSERT(m_nNumberOfPages      != 0);
	ATLASSERT(m_nMaxPageSize        >  0);
	ATLASSERT(m_spElementCollection != NULL);
	ATLASSERT(m_pages.size() == m_nNumberOfPages);

	// Find the page number.
	LONG nPageNumber = nIndex / m_nMaxPageSize;
	ATLASSERT((0 <= nPageNumber) && (nPageNumber < static_cast<LONG>(m_pages.size())));

	ElementPage& currentPage = m_pages[nPageNumber];
	if (NULL == currentPage.m_psaPageElements)
	{
		ATLASSERT(0 == currentPage.m_nPageSize);
		ATLASSERT(m_spElementCollection != NULL);
		ATLASSERT(currentPage.m_elements.size() == 0);

		SAFEARRAY* psa  = NULL;
		HRESULT    hRes = m_spElementCollection->GetPage(nPageNumber, &psa);
		if (FAILED(hRes) || (NULL == psa))
		{
			traceLog << "m_spElementCollection->GetPage failed with code " << hRes << " in CElementList::GetElement\n";
			return CComQIPtr<IElement>();
		}

		// UBound is the index of last element which might be zero for an array with just one element.
		hRes = ::SafeArrayGetUBound(psa, 1, &currentPage.m_nPageSize);
		if (FAILED(hRes))
		{
			traceLog << "SafeArrayGetUBound failed with code " << hRes << " in CElementList::GetElement\n";
			return CComQIPtr<IElement>();
		}

		currentPage.m_nPageSize++;
		if (currentPage.m_nPageSize <= 0)
		{
			traceLog << "Invalid ubound in CElementList::GetElement\n";
			return CComQIPtr<IElement>();			
		}

		currentPage.m_psaPageElements = psa;
		currentPage.m_elements.resize(currentPage.m_nPageSize);
	}

	LONG nElementIndexInPage = nIndex % m_nMaxPageSize;
	ATLASSERT((0 <= nElementIndexInPage) && (nElementIndexInPage < static_cast<LONG>(currentPage.m_elements.size())));
	ATLASSERT(currentPage.m_psaPageElements != NULL);

	if (currentPage.m_elements[nElementIndexInPage] != NULL)
	{
		return currentPage.m_elements[nElementIndexInPage];
	}
	else
	{
		ATLASSERT(nElementIndexInPage < currentPage.m_nPageSize);

		// I assume SafeArrayGetElement calls AddRef on &spDisp.
		CComQIPtr<IDispatch> spDisp;
		HRESULT hRes = ::SafeArrayGetElement(currentPage.m_psaPageElements, &nElementIndexInPage, &spDisp);
		if (FAILED(hRes) || (spDisp == NULL))
		{
			traceLog << "SafeArrayGetElement failed in CElementList::GetElement\n";
			return CComQIPtr<IElement>();
		}

		CComQIPtr<IElement> spResultElement;
		CreateElementObject(spDisp, &spResultElement);
		return (currentPage.m_elements[nElementIndexInPage] = spResultElement);
	}
}


STDMETHODIMP CElementList::get_core(ICore** pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CElementList::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in CElementList::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	CComQIPtr<ICore> spCore = m_spCore;
	*pVal = spCore.Detach();

	return HRES_OK;
}
