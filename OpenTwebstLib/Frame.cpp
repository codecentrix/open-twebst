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
#include "CodeErrors.h"
#include "HtmlHelpIDs.h"
#include "FindInTimeout.h"
#include "Core.h"
#include "Frame.h"
#include "HtmlHelpers.h"
#include "SearchContext.h"
#include "MethodAndPropertyNames.h"
#include "Browser.h"
#include "SearchCondition.h"

// CFrame

STDMETHODIMP CFrame::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IFrame
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}


void CFrame::SetHtmlWindow(IHTMLWindow2* pHtmlWindow)
{
	ATLASSERT(pHtmlWindow    != NULL);
	ATLASSERT(m_spHtmlWindow == NULL);
	m_spHtmlWindow = pHtmlWindow;
}


STDMETHODIMP CFrame::get_core(ICore** pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CFrame::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in CFrame::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	CComQIPtr<ICore> spCore = m_spCore;
	*pVal = spCore.Detach();

	return HRES_OK;
}


STDMETHODIMP CFrame::get_nativeFrame(IHTMLWindow2** pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CFrame::get_nativeFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_FRAME_PROPERTY, IDH_FRAME_NATIVE_FRAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	CComQIPtr<IHTMLWindow2> spWindow = m_spHtmlWindow;
	if (spWindow == NULL)
	{
		traceLog << "spWindow is NULL in CFrame::get_nativeFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_FRAME_PROPERTY, IDH_FRAME_NATIVE_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	*pVal = spWindow.Detach();
	return HRES_OK;
}


STDMETHODIMP CFrame::get_document(IHTMLDocument2** pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((NULL == pVal))
	{
		traceLog << "pVal is NULL in CFrame::get_document\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, DOCUMENT_PROPERTY, IDH_FRAME_DOCUMENT_PROPERTY);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, DOCUMENT_PROPERTY, IDH_FRAME_DOCUMENT_PROPERTY))
	{
		traceLog << "Invalid state in CFrame::get_document\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->GetDocumentFromWindow(m_spHtmlWindow, pVal);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling m_spPlugin->GetDocumentFromWindow in CFrame::get_document\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_FRAME_DOCUMENT_PROPERTY);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}
	else if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->GetDocumentFromWindow failed with code:" << hRes << " in CFrame::get_document\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, DOCUMENT_PROPERTY, IDH_FRAME_DOCUMENT_PROPERTY);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CFrame::FindElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CFrame::FindElement\n";
		SetComErrorMessage(IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDH_FRAME_FIND_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY;
	SearchContext sc("CFrame::FindElement", nSearchFlags, IDH_FRAME_FIND_ELEMENT,
	                 IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_ELEMENT_FAILED,
                     IDS_FIND_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CFrame::FindElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_FRAME_FIND_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::FindChildElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CFrame::FindChildElement\n";
		SetComErrorMessage(IDS_ERR_FIND_CHILD_ELEMENT_INVALID_PARAM, IDH_FRAME_FIND_CHILD_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_CHILDREN_ONLY;
	SearchContext sc("CFrame::FindChildElement", nSearchFlags, IDH_FRAME_FIND_CHILD_ELEMENT,
					 IDS_ERR_FIND_CHILD_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_CHILD_ELEMENT_FAILED,
					 IDS_FIND_CHILD_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CFrame::FindChildElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_FRAME_FIND_CHILD_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::FindAllElements(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppElementList)
	{
		traceLog << "ppElementList is null in CFrame::FindAllElements\n";
		SetComErrorMessage(IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDH_FRAME_FIND_ALL_ELEMENTS);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY | SEARCH_COLLECTION;
	SearchContext sc("CFrame::FindAllElements", nSearchFlags, IDH_FRAME_FIND_ALL_ELEMENTS,
	                 IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDS_ERR_FIND_ALL_ELEMENTS_FAILED,
                     IDS_FIND_ALL_ELEMENTS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElemListObject(spResult, ppElementList);
		if (NULL == *ppElementList)
		{
			traceLog << "Can not create an ElementList object in CFrame::FindAllElements\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_FRAME_FIND_ALL_ELEMENTS);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::FindChildrenElements(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppElementList)
	{
		traceLog << "ppElementList is null in CFrame::FindChildrenElements\n";
		SetComErrorMessage(IDS_ERR_FIND_CHILDREN_ELEMENTS_INVALID_PARAM, IDH_FRAME_FIND_CHILDREN_ELEM);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_CHILDREN_ONLY | SEARCH_COLLECTION;
	SearchContext sc("CFrame::FindChildrenElements", nSearchFlags, IDH_FRAME_FIND_CHILDREN_ELEM,
					 IDS_ERR_FIND_CHILDREN_ELEMENTS_INVALID_PARAM, IDS_ERR_FIND_CHILDREN_ELEMENTS_FAILED,
					 IDS_FIND_CHILDREN_ELEMENTS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElemListObject(spResult, ppElementList);
		if (NULL == *ppElementList)
		{
			traceLog << "Can not create an ElementList object in CFrame::FindChildrenElements\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_FRAME_FIND_CHILDREN_ELEM);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::FindFrame(BSTR bstrCond, IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppFrame)
	{
		traceLog << "ppFrame is null in CFrame::FindFrame\n";
		SetComErrorMessage(IDS_ERR_FIND_FRAME_INVALID_PARAM, IDH_FRAME_FIND_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_FRAME | SEARCH_ALL_HIERARCHY;
	SearchContext sc("CFrame::FindFrame", nSearchFlags, IDH_FRAME_FIND_FRAME,
	                 IDS_ERR_FIND_FRAME_INVALID_PARAM, IDS_ERR_FIND_FRAME_FAILED,
                     IDS_FIND_FRAME_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, NULL, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateFrameObject(spResult, ppFrame);
		if (NULL == *ppFrame)
		{
			traceLog << "Can not create a Frame object in CFrame::FindFrame\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_FRAME_FIND_FRAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::FindChildFrame(BSTR bstrCond, IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlWindow != NULL);
	if (NULL == ppFrame)
	{
		traceLog << "ppFrame is null in CFrame::FindChildFrame\n";
		SetComErrorMessage(IDS_ERR_FIND_CHILD_FRAME_INVALID_PARAM, IDH_FRAME_FIND_CHILD_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_FRAME | SEARCH_CHILDREN_ONLY;
	SearchContext sc("CFrame::FindChildFrame", nSearchFlags, IDH_FRAME_FIND_CHILD_FRAME,
					 IDS_ERR_FIND_CHILD_FRAME_INVALID_PARAM, IDS_ERR_FIND_CHILD_FRAME_FAILED,
					 IDS_FIND_CHILD_FRAME_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlWindow, sc, NULL, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateFrameObject(spResult, ppFrame);
		if (NULL == *ppFrame)
		{
			traceLog << "Can not create a Frame object in CFrame::FindChildFrame\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_FRAME_FIND_CHILD_FRAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CFrame::get_parentBrowser(IBrowser** ppBrowser)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppBrowser)
	{
		traceLog << "ppBrowser parameter is NULL in CFrame::get_parentBrowser\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_BROWSER_PROPERTY, IDH_FRAME_PARENT_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PARENT_BROWSER_PROPERTY, IDH_FRAME_PARENT_BROWSER))
	{
		traceLog << "Invalid state in CFrame::get_parentBrowser\n";
		return HRES_FAIL;
	}

	// Create a new frame object.
	IBrowser* pNewBrowser = NULL;
	HRESULT   hRes      = CComCoClass<CBrowser>::CreateInstance(&pNewBrowser);
	if (pNewBrowser != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CBrowser* pBrowserObject = static_cast<CBrowser*>(pNewBrowser); // Down cast !!!
		ATLASSERT(pBrowserObject != NULL);
		pBrowserObject->SetPlugin(m_spPlugin);
		pBrowserObject->SetCore(m_spCore);

		// Set the object reference into the output pointer.
		*ppBrowser = pNewBrowser;
	}
	else
	{
		traceLog << "Can not create a Browser object in CFrame::get_parentBrowser\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_BROWSER, IDH_FRAME_PARENT_BROWSER);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CFrame::get_parentFrame(IFrame** ppParentFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppParentFrame)
	{
		traceLog << "ppParentFrame is NULL in CBrowser::get_parentFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_FRAME_PARENT_FRAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_FRAME_PARENT_FRAME))
	{
		traceLog << "Invalid state in CFrame::get_parentFrame\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLWindow2> spParentFrame;
	HRESULT hRes = m_spHtmlWindow->get_parent(&spParentFrame);
	if (FAILED(hRes))
	{
		traceLog << "IHTTMLWindow2::get_parent failed in CFrame::get_parentFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_FRAME_PARENT_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Create an IFrame object.
	if ((spParentFrame != NULL) && (!spParentFrame.IsEqualObject(m_spHtmlWindow)))
	{
		hRes = CreateFrameObject(spParentFrame, ppParentFrame);
		if (NULL == *ppParentFrame)
		{
			ATLASSERT(FAILED(hRes));

			traceLog << "Can not create a Frame object in CFrame::get_parentFrame\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_FRAME_PARENT_FRAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}
	else
	{
		*ppParentFrame = NULL;
	}

	return HRES_OK;
}


BOOL CFrame::IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID)
{
	if (!BaseLibObject::IsValidState(nPatternID, szMethodName, dwHelpID))
	{
		traceLog << "IsValidState CFrame::IsValidState\n";
		return FALSE;
	}
	else
	{
		if (m_spHtmlWindow == NULL)
		{
			traceLog << "m_spHtmlWindow is NULL in CFrame::IsValidState\n";
			SetComErrorMessage(nPatternID, szMethodName, dwHelpID);
			SetLastErrorCode(ERR_FAIL);
			return FALSE;
		}
	}

	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////////////////
// Methods and properties not yet implemented.
STDMETHODIMP CFrame::get_frameElement(IElement** ppFrameElement)
{
	ATLASSERT(ppFrameElement != NULL);
	return E_NOTIMPL;
}


STDMETHODIMP CFrame::get_title(BSTR* pVal)
{
	ATLASSERT(pVal != NULL);
	return E_NOTIMPL;
}


STDMETHODIMP CFrame::get_url(BSTR* pVal)
{
	ATLASSERT(pVal != NULL);
	return E_NOTIMPL;
}
