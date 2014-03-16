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
#include "CodeErrors.h"
#include "HtmlHelpIDs.h"
#include "Registry.h"
#include "Settings.h"
#include "DebugServices.h"
#include "Core.h"
#include "Element.h"
#include "FindInTimeout.h"
#include "SearchContext.h"
#include "..\OTWBSTInjector\OTWBSTInjector.h"
#include "MethodAndPropertyNames.h"
#include "Browser.h"
#include "Keyboard.h"
#include "HtmlHelpers.h"
#include "SearchCondition.h"
using namespace Common;



struct SelectCallContext
{
	SelectCallContext(const TCHAR* szFunctionName, ULONG nSearchFlags, DWORD  dwHelpContextID) :
		m_szFunctionName(szFunctionName), m_nSearchFlags(nSearchFlags), m_dwHelpContextID(dwHelpContextID)
	{
		ATLASSERT(szFunctionName != NULL);
	}

	const TCHAR* m_szFunctionName;
	ULONG        m_nSearchFlags;
	DWORD        m_dwHelpContextID;
};


struct ClickCallContext
{
	ClickCallContext(BOOL bRightClick, const TCHAR* szFunctionName, DWORD  dwHelpContextID, LONG x = -1, LONG y = -1) :
		m_bRightClick(bRightClick), m_szFunctionName(szFunctionName),  m_dwHelpContextID(dwHelpContextID),
		m_nClickX(x), m_nClickY(y)
	{
		ATLASSERT(szFunctionName != NULL);
	}

	BOOL         m_bRightClick;
	const TCHAR* m_szFunctionName;
	DWORD        m_dwHelpContextID;
	LONG         m_nClickX;
	LONG         m_nClickY;
};


STDMETHODIMP CElement::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IElement
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}


void CElement::SetHtmlElement(IHTMLElement* pHtmlElement)
{
	ATLASSERT(pHtmlElement    != NULL);
	ATLASSERT(m_spHtmlElement == NULL);
	m_spHtmlElement = pHtmlElement;
}

STDMETHODIMP CElement::get_core(ICore** pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CElement::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_ELEMENT_CORE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in CElement::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_ELEMENT_CORE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	CComQIPtr<ICore> spCore = m_spCore;
	*pVal = spCore.Detach();

	return HRES_OK;
}


STDMETHODIMP CElement::get_nativeElement(IHTMLElement** pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CElement::get_nativeElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_ELEMENT_PROPERTY, IDH_ELEMENT_NATIVE_ELEMENT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	CComQIPtr<IHTMLElement> spElement = m_spHtmlElement;
	if (spElement == NULL)
	{
		traceLog << "spElement is NULL in CElement::get_nativeElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_ELEMENT_PROPERTY, IDH_ELEMENT_NATIVE_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	*pVal = spElement.Detach();
	return HRES_OK;
}


STDMETHODIMP CElement::FindElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlElement != NULL);
	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CElement::FindElement\n";
		SetComErrorMessage(IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDH_ELEMENT_FIND_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY;
	SearchContext sc("CElement::FindElement", nSearchFlags, IDH_ELEMENT_FIND_ELEMENT,
	                 IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_ELEMENT_FAILED,
                     IDS_FIND_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CElement::FindElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_FIND_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::FindChildElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlElement != NULL);
	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CElement::FindChildElement\n";
		SetComErrorMessage(IDS_ERR_FIND_CHILD_ELEMENT_INVALID_PARAM, IDH_ELEMENT_FIND_CHILD_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_CHILDREN_ONLY;
	SearchContext sc("CElement::FindChildElement", nSearchFlags, IDH_ELEMENT_FIND_CHILD_ELEMENT,
					 IDS_ERR_FIND_CHILD_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_CHILD_ELEMENT_FAILED,
					 IDS_FIND_CHILD_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CElement::FindChildElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_FIND_CHILD_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::FindAllElements(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlElement != NULL);
	if (NULL == ppElementList)
	{
		traceLog << "ppElementList is null in CElement::FindAllElements\n";
		SetComErrorMessage(IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDH_ELEMENT_FIND_ALL_ELEMENTS);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY | SEARCH_COLLECTION;
	SearchContext sc("CElement::FindAllElements", nSearchFlags, IDH_ELEMENT_FIND_ALL_ELEMENTS,
	                 IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDS_ERR_FIND_ALL_ELEMENTS_FAILED,
					 IDS_FIND_ALL_ELEMENTS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElemListObject(spResult, ppElementList);
		if (NULL == *ppElementList)
		{
			traceLog << "Can not create an ElementList object in CElement::FindAllElements\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_ELEMENT_FIND_ALL_ELEMENTS);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::FindChildrenElements(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlElement != NULL);
	if (NULL == ppElementList)
	{
		traceLog << "ppElementList is null in CElement::FindChildrenElements\n";
		SetComErrorMessage(IDS_ERR_FIND_CHILDREN_ELEMENTS_INVALID_PARAM, IDH_ELEMENT_FIND_CHILDREN_ELEM);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_CHILDREN_ONLY | SEARCH_COLLECTION;
	SearchContext sc("CElement::FindChildrenElements", nSearchFlags, IDH_ELEMENT_FIND_CHILDREN_ELEM,
					 IDS_ERR_FIND_CHILDREN_ELEMENTS_INVALID_PARAM, IDS_ERR_FIND_CHILDREN_ELEMENTS_FAILED,
					 IDS_FIND_CHILDREN_ELEMENTS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElemListObject(spResult, ppElementList);
		if (NULL == *ppElementList)
		{
			traceLog << "Can not create an ElementList object in CElement::FindChildrenElements\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_ELEMENT_FIND_CHILDREN_ELEM);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::Click()
{
	FIRE_CANCEL_REQUEST();

	ClickCallContext clickCtx(FALSE, CLICK_METHOD, IDH_ELEMENT_CLICK);
	return Click(clickCtx, TRUE);
}


STDMETHODIMP CElement::RightClick()
{
	FIRE_CANCEL_REQUEST();

	ClickCallContext clickCtx(TRUE, RIGHT_CLICK_METHOD, IDH_ELEMENT_RIGHT_CLICK);
	return Click(clickCtx, TRUE);
}


HRESULT CElement::Click(const ClickCallContext& clickCtx, BOOL bClickFileInputButton, BOOL* pbPostedClick)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (pbPostedClick != NULL)
	{
		*pbPostedClick = FALSE;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID))
	{
		traceLog << "Invalid state in CElement::Click\n";
		return HRES_FAIL;
	}

	// Get the useRegExp Core property value.
	VARIANT_BOOL   vbUseHardwareEvents;
	HRESULT hRes = m_spCore->get_useHardwareInputEvents(&vbUseHardwareEvents);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_useIEevents failed in CElement::Click with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (VARIANT_FALSE == vbUseHardwareEvents)
	{
		VARIANT_BOOL vbClickAsync = VARIANT_FALSE;
		hRes = m_spCore->get_asyncHtmlEvents(&vbClickAsync);

		if (FAILED(hRes))
		{
			traceLog << "ICore::get_asyncHtmlEvents failed in CElement::Click with code " << hRes << "\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}

		if (VARIANT_FALSE == vbClickAsync)
		{
			// Use IE click events.
			hRes = m_spPlugin->ClickElementAt(m_spHtmlElement, -1, -1, clickCtx.m_bRightClick);
		}
		else
		{
			hRes = m_spPlugin->ClickElementAtAsync(m_spHtmlElement, -1, -1, clickCtx.m_bRightClick);
		}

		if (FAILED(hRes))
		{
			traceLog << "m_spHtmlElement->click failed in CElement::Click with code " << hRes << "\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}
	else // Use hardware event.
	{
		// Get the web browser object.
		CComQIPtr<IWebBrowser2> spWebBrws;
		HRESULT hRes = m_spPlugin->GetNativeBrowser(&spWebBrws);

		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			ATLASSERT(spWebBrws == NULL);
			traceLog << "Connection with the browser lost while calling m_spPlugin->GetNativeBrowser in CElement::Click\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, clickCtx.m_dwHelpContextID);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}

		if (spWebBrws == NULL)
		{
			ATLASSERT(FAILED(hRes));

			traceLog << "m_spPlugin->GetNativeBrowser failed with code " << hRes << "in CElement::Click\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}

		HWND hTabWnd   = NULL;
		HWND hIESrvWnd = HtmlHelpers::GetIEWndFromBrowser(spWebBrws);

		if (Common::GetIEVersion() > 6)
		{
			hTabWnd = hIESrvWnd;

			if (!::IsWindow(hTabWnd))
			{
				traceLog << "HtmlHelpers::GetIEWndFromBrowser failed in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);

				return HRES_FAIL;
			}

			ATLASSERT((Common::GetWndClass(hTabWnd) == _T("TabWindowClass")) ||
					  (Common::GetWndClass(hTabWnd) == _T("Shell Embedding")));
		}

		// Get the window handle of the IE instance.
		LONG_PTR nIeWnd = NULL;
		hRes = spWebBrws->get_HWND(&nIeWnd);

		HWND hIeWnd = (HWND)nIeWnd;
		if (FAILED(hRes) || !::IsWindow(hIeWnd))
		{
			// Maybe it is an embeded browser control.
			hIeWnd = Common::GetTopParentWnd(hIESrvWnd);

			if (!::IsWindow(hIeWnd))
			{
				traceLog << "IWebBrowser2::get_HWND failed with code " << hRes << "in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);

				return HRES_FAIL;
			}
		}

		// The class window can be anything if it is an embeded browser control.
		//ATLASSERT(Common::GetWndClass(hIeWnd) == _T("IEFrame"));

		// If the tab is not the active tab then hTabWnd is disabled.
		// If the whole browser is hidden then hTabWnd is not visible.
		// If hTabWnd is NULL then it should be IE6.
		BOOL bIsTabVisible       = ((hTabWnd != NULL) && ::IsWindow(hTabWnd) && ::IsWindowEnabled(hTabWnd) && ::IsWindowVisible(hTabWnd));
		BOOL bIsIE6TopWndVisible = ((NULL == hTabWnd) && ::IsWindowVisible(hIeWnd));
		BOOL bUseMouseEvents     = bIsTabVisible || bIsIE6TopWndVisible;

		if (bUseMouseEvents)
		{
			// Put the IE window in foreground.
			BOOL bRes = PutWindowInForeground(hIeWnd);
			if (!bRes)
			{
				traceLog << "PutWindowInForeground failed in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			// Get the screen click point.
			LONG nScreenX;
			LONG nScreenY;
			hRes = m_spPlugin->GetScreenClickPoint(m_spHtmlElement, bClickFileInputButton, &nScreenX, &nScreenY);
			if (FAILED(hRes))
			{
				traceLog << "m_spPlugin->GetScreenClickPoint failed with code " << hRes << "in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			nScreenX += 2;
			nScreenY += 2;

			if (!clickCtx.m_bRightClick)
			{
				// Generate a mouse click hardware event.
				::SetCursorPos(nScreenX, nScreenY);
				mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN,  (DWORD)nScreenX, (DWORD)nScreenY, 0, HARDWARE_EVENT_EXTRA_INFO);
				mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTUP,    (DWORD)nScreenX, (DWORD)nScreenY, 0, HARDWARE_EVENT_EXTRA_INFO);
				this->Sleep(INTERNAL_GLOBAL_PAUSE);
			}
			else
			{
				// Generate a mouse click hardware event.
				::SetCursorPos(nScreenX, nScreenY);
				mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_RIGHTDOWN, (DWORD)nScreenX, (DWORD)nScreenY, 0, HARDWARE_EVENT_EXTRA_INFO);
				mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_RIGHTUP,   (DWORD)nScreenX, (DWORD)nScreenY, 0, HARDWARE_EVENT_EXTRA_INFO);
				this->Sleep(INTERNAL_GLOBAL_PAUSE);
			}
		}
		else
		{
			HWND hIEServWnd = NULL;

			// Get "Internet Explorer_Server" window.
			if (Common::GetIEVersion() > 6)
			{
				ATLASSERT(::IsWindow(hTabWnd));
				ATLASSERT(Common::GetWndClass(hTabWnd) == _T("TabWindowClass"));
				hIEServWnd = Common::GetChildWindowByClassName(hTabWnd, _T("Internet Explorer_Server"));
			}
			else
			{
				// IE6.
				ATLASSERT(!::IsWindowVisible(hIeWnd));
				ATLASSERT(Common::GetWndClass(hIeWnd) == _T("IEFrame"));

				LONG nWnd    = 0;
				hRes         = m_spPlugin->GetIEServerWnd(m_spHtmlElement, &nWnd);
				hIEServWnd = static_cast<HWND>(LongToHandle(nWnd));
			}

			if (NULL == hIEServWnd)
			{
				traceLog << "Common::GetChildWindowByClassName failed in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			ATLASSERT(Common::GetWndClass(hIEServWnd) == _T("Internet Explorer_Server"));

			// Get the screen click point.
			LONG nScreenX;
			LONG nScreenY;
			hRes = m_spPlugin->GetScreenClickPoint(m_spHtmlElement, bClickFileInputButton, &nScreenX, &nScreenY);
			if (FAILED(hRes))
			{
				traceLog << "m_spPlugin->GetScreenClickPoint failed with code " << hRes << "in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			POINT clientClickPoint = { nScreenX, nScreenY };
			BOOL  bConvert         = ::ScreenToClient(hIEServWnd, &clientClickPoint);

			if (!bConvert)
			{
				traceLog << "ScreenToClient failed with in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			UINT nMouseDownMsg = 0;
			UINT nMouseUpMsg   = 0;

			if (!clickCtx.m_bRightClick)
			{
				nMouseDownMsg = WM_LBUTTONDOWN;
				nMouseUpMsg   = WM_LBUTTONUP;
			}
			else
			{
				nMouseDownMsg = WM_RBUTTONDOWN;
				nMouseUpMsg   = WM_RBUTTONUP;
			}

			BOOL bPost = ::PostMessage(hIEServWnd, nMouseDownMsg, (WPARAM)MK_LBUTTON, MAKELPARAM(clientClickPoint.x, clientClickPoint.y));
			if (!bPost)
			{
				traceLog << "PostMessage(WM_LBUTTONDOWN) failed with in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			bPost = ::PostMessage(hIEServWnd, nMouseUpMsg, (WPARAM)0, MAKELPARAM(clientClickPoint.x, clientClickPoint.y));
			if (!bPost)
			{
				traceLog << "PostMessage(WM_LBUTTONUP) failed with in CElement::Click\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, clickCtx.m_szFunctionName, clickCtx.m_dwHelpContextID);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			if (pbPostedClick != NULL)
			{
				*pbPostedClick = TRUE;
			}
		}
	}

	return HRES_OK;
}


STDMETHODIMP CElement::get_parentElement(IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppElement)
	{
		traceLog << "pVal parameter is NULL in CElement::get_parentElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_ELEMENT_PROPERTY, IDH_ELEMENT_PARENT_ELEMENT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PARENT_ELEMENT_PROPERTY, IDH_ELEMENT_PARENT_ELEMENT))
	{
		traceLog << "Invalid state in CElement::get_parentElement\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLElement> spParentElement;
	HRESULT hRes = m_spHtmlElement->get_parentElement(&spParentElement);

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling IHTMLElement::get_parentElement in CElement::get_parentElement\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_PARENT_ELEMENT);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "IHTMLElement::get_parentElement failed with code " << hRes << " in CElement::get_parentElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_ELEMENT_PROPERTY, IDH_ELEMENT_PARENT_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (spParentElement == NULL)
	{
		*ppElement = NULL;
		return HRES_OK;
	}

	HRESULT hCreateElem = CreateElementObject(spParentElement, ppElement);
	if (NULL == *ppElement)
	{
		traceLog << "Can not create an Element object in CElement::get_parentElement\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_PARENT_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CElement::get_nextSiblingElement(IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppElement)
	{
		traceLog << "pVal parameter is NULL in CElement::get_nextSiblingElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NEXT_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_NEXT_SIBLING);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, NEXT_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_NEXT_SIBLING))
	{
		traceLog << "Invalid state in CElement::get_nextSiblingElement\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLDOMNode> spNode;
	CComQIPtr<IHTMLElement> spNextElement;

	HRESULT hRes = m_spHtmlElement->QueryInterface(&spNode);
	if (spNode != NULL)
	{
		while (TRUE)
		{
			CComQIPtr<IHTMLDOMNode> spNextNode;
			hRes = spNode->get_nextSibling(&spNextNode);
			if (spNextNode != NULL)
			{
				hRes = spNextNode->QueryInterface(&spNextElement);
				if (spNextElement != NULL)
				{
					break;
				}
				else if (E_NOINTERFACE != hRes)
				{
					break;
				}

				spNode = spNextNode;
			}
			else
			{
				break;
			}
		}
	}

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost in CElement::get_nextSiblingElement\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_NEXT_SIBLING);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	// E_NOINTERFACE when the next node is not an HTML element.
	if (FAILED(hRes) && (hRes != E_NOINTERFACE))
	{
		traceLog << "Fail to get the next element, code " << hRes << " in CElement::get_nextSiblingElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NEXT_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_NEXT_SIBLING);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (spNextElement == NULL)
	{
		*ppElement = NULL;
		return HRES_OK;
	}

	HRESULT hCreateElem = CreateElementObject(spNextElement, ppElement);
	if (NULL == *ppElement)
	{
		traceLog << "Can not create an Element object in CElement::get_nextSiblingElement\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_NEXT_SIBLING);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CElement::get_previousSiblingElement(IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppElement)
	{
		traceLog << "pVal parameter is NULL in CElement::get_previousSiblingElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PREVIOUS_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_PREVIOUS_SIBLING);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PREVIOUS_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_PREVIOUS_SIBLING))
	{
		traceLog << "Invalid state in CElement::get_previousSiblingElement\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLDOMNode> spNode;
	CComQIPtr<IHTMLElement> spPrevisouElement;

	HRESULT hRes = m_spHtmlElement->QueryInterface(&spNode);
	if (spNode != NULL)
	{
		while (TRUE)
		{
			CComQIPtr<IHTMLDOMNode> spPrevisouNode;
			hRes = spNode->get_previousSibling(&spPrevisouNode);
			if (spPrevisouNode != NULL)
			{
				hRes = spPrevisouNode->QueryInterface(&spPrevisouElement);
				if (spPrevisouElement != NULL)
				{
					break;
				}
				else if (E_NOINTERFACE != hRes)
				{
					break;
				}
			}
			else
			{
				break;
			}
		}
	}

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost in CElement::get_previousSiblingElement\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_PREVIOUS_SIBLING);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	// E_NOINTERFACE when the next node is not an HTML element.
	if (FAILED(hRes) && (hRes != E_NOINTERFACE))
	{
		traceLog << "Fail to get the previsou element, code " << hRes << " in CElement::get_previousSiblingElement\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PREVIOUS_SIBLING_ELEM_PROPERTY, IDH_ELEMENT_PREVIOUS_SIBLING);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (spPrevisouElement == NULL)
	{
		*ppElement = NULL;
		return HRES_OK;
	}

	HRESULT hCreateElem = CreateElementObject(spPrevisouElement, ppElement);
	if (NULL == *ppElement)
	{
		traceLog << "Can not create an Element object in CElement::get_previousSiblingElement\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_PREVIOUS_SIBLING);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CElement::get_parentFrame(IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppFrame)
	{
		traceLog << "ppFrame parameter is NULL in CElement::get_parentFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_ELEMENT_PARENT_FRAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_ELEMENT_PARENT_FRAME))
	{
		traceLog << "Invalid state in CElement::get_parentFrame\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLWindow2> spParentWindow;
	HRESULT hRes = GetParentWindow(m_spHtmlElement, &spParentWindow);

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost in CElement::get_parentFrame\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_PARENT_FRAME);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "Fail to get the previsou element, code " << hRes << " in CElement::get_parentFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_FRAME_PROPERTY, IDH_ELEMENT_PARENT_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	ATLASSERT(spParentWindow != NULL);
	hRes = CreateFrameObject(spParentWindow, ppFrame);
	if (NULL == *ppFrame)
	{
		traceLog << "Can not create a new Frame object in CElement::get_parentFrame\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_ELEMENT_PARENT_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


HRESULT CElement::GetParentWindow(IHTMLElement* pElement, IHTMLWindow2** ppElement)
{
	ATLASSERT(pElement  != NULL);
	ATLASSERT(ppElement != NULL);
	ATLASSERT(NULL == *ppElement);

	// Get the document of the element.
	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = pElement->get_document(&spDisp);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLElement::get_document failed with code " << hRes << " in CElement::GetParentWindow\n";
		return hRes;
	}

	CComQIPtr<IHTMLDocument2> spDocument;
	hRes = spDisp->QueryInterface(&spDocument);
	if (FAILED(hRes))
	{
		traceLog << "Query for IHTMLDocument2 failed with code " << hRes << " in CElement::GetParentWindow\n";
		return hRes;
	}

	// Get the window object.
	hRes = spDocument->get_parentWindow(ppElement);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLDocument2::get_parentWindow failed with code " << hRes << " in CElement::GetParentWindow\n";
	}

	return hRes;
}


STDMETHODIMP CElement::get_parentBrowser(IBrowser** ppBrowser)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppBrowser)
	{
		traceLog << "ppBrowser parameter is NULL in CElement::get_parentBrowser\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, PARENT_BROWSER_PROPERTY, IDH_ELEMENT_PARENT_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, PARENT_BROWSER_PROPERTY, IDH_ELEMENT_PARENT_BROWSER))
	{
		traceLog << "Invalid state in CElement::get_parentBrowser\n";
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
		traceLog << "Can not create a Browser object in CElement::get_parentBrowser\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_BROWSER, IDH_ELEMENT_PARENT_BROWSER);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CElement::InputText(BSTR bstrText)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// Check parameter.
	if (bstrText == NULL)
	{
		traceLog << "bstrText parameter is NULL in CElement::InputText\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT))
	{
		traceLog << "Invalid state in CElement::InputText\n";
		return HRES_FAIL;
	}

	HTML_EDIT_BOX_TYPE editBoxType = HTML_NOT_EDIT_BOX;
	HRESULT hRes = IsHtmlEditBox(m_spHtmlElement, &editBoxType);
	if (S_FALSE == hRes)
	{
		traceLog << "InputText is not applicable to the current html element in CElement::InputText\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, IDH_ELEMET_INPUT_TEXT);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);
		return HRES_OPERATION_NOT_APPLICABLE;
	}
	else if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost while calling IsHtmlEditBox in CElement::InputText\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "IsHtmlEditBox failed in CElement::InputText\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	// Get the useRegExp Core property value.
	VARIANT_BOOL vbUseHardwareEvents;
	hRes = m_spCore->get_useHardwareInputEvents(&vbUseHardwareEvents);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_useIEevents failed in CElement::InputText with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	int nIeVer = Common::GetIEVersion();
	if ((HTML_FILE_EDIT_BOX == editBoxType) && (nIeVer > 7))
	{
		// On IE8 <input type="file" /> is read-only.
		return InputTextInFileCtrlIE8(bstrText);
	}
	else
	{
		if (VARIANT_FALSE == vbUseHardwareEvents)
		{
			// Get the asyncHtmlEvents Core flag.
			VARIANT_BOOL vbAsyncEvents;
			hRes = m_spCore->get_asyncHtmlEvents(&vbAsyncEvents);
			if (FAILED(hRes))
			{
				traceLog << "ICore::get_asyncHtmlEvents failed in CElement::InputText with code " << hRes << "\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}

			// Use IE events.
			return IeInputText(bstrText, (HTML_FILE_EDIT_BOX == editBoxType), vbAsyncEvents);
		}
		else
		{
			// Use hardware events.
			return HardwareInputText(bstrText, (HTML_FILE_EDIT_BOX == editBoxType));
		}
	}
}


BOOL CElement::GetElemValue(CComQIPtr<IHTMLElement> spHtmlElement, BSTR* pBstrInitialValue)
{
	ATLASSERT(spHtmlElement     != NULL);
	ATLASSERT(pBstrInitialValue != NULL);
	ATLASSERT(NULL == *pBstrInitialValue);

	CComQIPtr<IHTMLInputElement> spInputElement = spHtmlElement;
	if (spInputElement != NULL)
	{
		HRESULT hRes = spInputElement->get_value(pBstrInitialValue);
		return SUCCEEDED(hRes);
	}

	CComQIPtr<IHTMLTextAreaElement> spTextAreaElement = spHtmlElement;
	if (spTextAreaElement != NULL)
	{
		HRESULT hRes = spTextAreaElement->get_value(pBstrInitialValue);
		return SUCCEEDED(hRes);
	}

	ATLASSERT(FALSE);
	return FALSE;
}


// In IE8 only the FrameTab parent of "Internet Explorer_Server" is disabled when the tab is not active.
BOOL CElement::IsWindowActive(HWND hIEWnd)
{
	if (!::IsWindow(hIEWnd))
	{
		return FALSE;
	}

	HWND hCrntWnd = hIEWnd;

	while (TRUE)
	{
		if (!::IsWindowEnabled(hCrntWnd))
		{
			return FALSE;
		}

		hCrntWnd = ::GetParent(hCrntWnd);

		if (!::IsWindow(hCrntWnd))
		{
			// Reached the top.
			break;
		}
	}

	return TRUE;
}


HRESULT CElement::InputTextInFileCtrlIE8(BSTR bstrText)
{
	ATLASSERT(bstrText        != NULL);
	ATLASSERT(m_spPlugin      != NULL);
	ATLASSERT(m_spHtmlElement != NULL);

	CPath pathToFile(bstrText);
	if (!pathToFile.FileExists())
	{
		traceLog << "Invalid file path in CElement::InputTextInFileCtrlIE8\n";
		SetComErrorMessage(IDS_INVALID_FILE_PATH_IN_INPUT_TEXT, IDH_ELEMET_INPUT_TEXT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	HWND    hIeServerWnd = NULL;
	LONG    nWnd         = 0;
	HRESULT hRes         = m_spPlugin->GetIEServerWnd(NULL, &nWnd);

	hIeServerWnd = static_cast<HWND>(LongToHandle(nWnd));
	if (SUCCEEDED(hRes) && ::IsWindow(hIeServerWnd))
	{
		if (!IsWindowActive(hIeServerWnd))
		{
			// The tab must be active.
			traceLog << "The tab is not active in CElement::InputTextInFileCtrlIE8\n";
			SetComErrorMessage(IDS_INPUT_TEXT_TAB_NOT_ACTIVE, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_FAIL);

			return HRES_FAIL;
		}
	}
	else
	{
		traceLog << "GetIEServerWnd failed in CElement::InputTextInFileCtrlIE8\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	// Get the value of the element.
	CComBSTR bstrInitialValue;
	BOOL bInitValue = GetElemValue(m_spHtmlElement, &bstrInitialValue);

	// Set the value if it is different from the current value.
	if (!bInitValue || (bInitValue && (bstrInitialValue != bstrText)))
	{
		// Click asynchronously the upload control so "Choose file" dialog box is displayed.
		HRESULT hRes = m_spPlugin->ClickElementAtAsync(m_spHtmlElement, -1, -1, FALSE);

		if (SUCCEEDED(hRes))
		{
			hRes = InputTextInFileDlg(bstrText, hIeServerWnd);
		}

		if (FAILED(hRes))
		{
			traceLog << "FAILED(hRes) in CElement::InputTextInFileCtrlIE8\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_FAIL);

			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


BOOL CElement::FindFileDlgCallback(HWND hWnd, void*)
{
	ATLASSERT(Common::GetWndClass(hWnd) == _T("#32770"));

	if (!::IsWindowVisible(hWnd) || !::IsWindowEnabled(hWnd))
	{
		return FALSE;
	}

	HWND hComboExWnd = Common::GetChildWindowByClassName(hWnd, _T("ComboBoxEx32"));
	if (::IsWindow(hComboExWnd))
	{
		// The dialog box contains a ComboBoxEx32 control. That must be it.
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


BOOL CElement::FindOpenBtnCallback(HWND hWnd, void*)
{
	ATLASSERT(Common::GetWndClass(hWnd) == _T("Button"));

	int nCtrlID = ::GetDlgCtrlID(hWnd);
	return (IDOK == nCtrlID);
}


HRESULT CElement::InputTextInFileDlg(BSTR bstrText, HWND hIeWnd)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(bstrText        != NULL);
	ATLASSERT(m_spPlugin      != NULL);
	ATLASSERT(m_spHtmlElement != NULL);

	HWND hChooseFileDlg = NULL;
	DWORD dwStartTime   = ::GetTickCount();

	while (TRUE)
	{
		FIRE_CANCEL_REQUEST();

		hChooseFileDlg = Common::GetTopLevelWindowByClassName(hIeWnd, _T("#32770"), FindFileDlgCallback);
		if (::IsWindow(hChooseFileDlg))
		{
			// We have found the window.
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > Common::INTERNAL_GLOBAL_TIMEOUT)
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		this->Sleep(Common::INTERNAL_GLOBAL_PAUSE);
	}

	if (::IsWindow(hChooseFileDlg))
	{
		return InputTextInFileDlg(hChooseFileDlg, bstrText);
	}
	else
	{
		traceLog << "hChooseFileDlg is not a window in CElement::InputTextInFileDlg\n";
		return HRES_FAIL;
	}
}


HRESULT CElement::InputTextInFileDlg(HWND hDlgWnd, BSTR bstrText)
{
	ATLASSERT(bstrText != NULL);
	ATLASSERT(Common::GetWndClass(hDlgWnd) == _T("#32770"));

	HWND hComboExWnd = Common::GetChildWindowByClassName(hDlgWnd, _T("ComboBoxEx32"));
	if (::IsWindow(hComboExWnd))
	{
		HWND hComboWnd = Common::GetChildWindowByClassName(hComboExWnd, _T("ComboBox"));
		if (::IsWindow(hComboWnd))
		{
			HWND hEditWnd = Common::GetChildWindowByClassName(hComboWnd, _T("Edit"));
			if (::IsWindow(hEditWnd))
			{
				USES_CONVERSION;
				::SendMessage(hEditWnd, WM_SETTEXT, 0, (LPARAM)W2T(bstrText));

				return PressOpenBtnInFileDlg(hDlgWnd);
			}
			else
			{
				traceLog << "hEditWnd is not a window in CElement::InputTextInFileDlg\n";
			}
		}
		else
		{
			traceLog << "hCombo is not a window in CElement::InputTextInFileDlg\n";
		}
	}
	else
	{
		traceLog << "hComboExWnd is not a window in CElement::InputTextInFileDlg\n";
	}

	return HRES_FAIL;
}


HRESULT CElement::PressOpenBtnInFileDlg(HWND hDlgWnd)
{
	ATLASSERT(Common::GetWndClass(hDlgWnd) == _T("#32770"));
	HWND hOpenBtnWnd = Common::GetChildWindowByClassName(hDlgWnd, _T("Button"), FindOpenBtnCallback);

	if (::IsWindow(hOpenBtnWnd))
	{
		HWND hDlgWnd = ::GetParent(hOpenBtnWnd);

		if (::IsWindow(hDlgWnd))
		{
			::SendMessage(hDlgWnd, WM_COMMAND, MAKEWPARAM(IDOK, BN_CLICKED), (LPARAM)hOpenBtnWnd);
			return HRES_OK;
		}
		else
		{
			traceLog << "hDlgWnd is not a window in CElement::PressOpenBtnInFileDlg\n";
			return HRES_FAIL;
		}
	}
	else
	{
		traceLog << "hOpenBtnWnd is not a window in CElement::PressOpenBtnInFileDlg\n";
		return HRES_FAIL;
	}
}


HWND CElement::GetTopWnd()
{
	ATLASSERT(m_spPlugin != NULL);

	HWND    hIeServerWnd = NULL;
	LONG    nWnd         = 0;
	HRESULT hRes         = m_spPlugin->GetIEServerWnd(NULL, &nWnd);

	hIeServerWnd = static_cast<HWND>(LongToHandle(nWnd));
	if (SUCCEEDED(hRes) && ::IsWindow(hIeServerWnd))
	{
		ATLASSERT(Common::GetWndClass(hIeServerWnd) == _T("Internet Explorer_Server"));

		// Get the top level IE window.
		HWND hTopWnd = Common::GetTopParentWnd(hIeServerWnd);

		return hTopWnd;
	}
	else
	{
		traceLog << "GetIEServerWnd failed in CElement::GetTopWnd\n";
		return NULL;
	}
}


HRESULT CElement::IeInputText(BSTR bstrText, BOOL bIsInputFileType, VARIANT_BOOL vbAsync)
{
	ATLASSERT(bstrText        != NULL);
	ATLASSERT(m_spPlugin      != NULL);
	ATLASSERT(m_spHtmlElement != NULL);

	// Set the focus on html element.
	HRESULT hFocus = m_spPlugin->SetFocusOnElement(m_spHtmlElement, vbAsync);
	if (FAILED(hFocus))
	{
		// We can live without focus. Continue.
		traceLog << "Failed to focus element in CElement::IeInputText\n";
	}

	// Get the value of the element.
	CComBSTR bstrInitialValue;
	BOOL bInitValue = GetElemValue(m_spHtmlElement, &bstrInitialValue);

	HRESULT  hRes = S_OK;
	if (!bIsInputFileType)
	{
		// Rise events for each character in input string.
		size_t   nLen             = wcslen(bstrText);
		BOOL     bRes             = TRUE;
		CComBSTR bstrCurrentValue = "";

		for (unsigned i = 0; i < nLen; ++i)
		{
			hRes = m_spPlugin->FireEventOnElement(m_spHtmlElement, CComBSTR("onkeydown"), bstrText[i], vbAsync);
			if (FAILED(hRes))
			{
				traceLog << "Failed to rise onkeydown event in CElement::IeInputText\n";
				break;
			}

			hRes = m_spPlugin->FireEventOnElement(m_spHtmlElement, CComBSTR("onkeypress"), bstrText[i], vbAsync);
			if (FAILED(hRes))
			{
				traceLog << "Failed to rise onkeypress event in CElement::IeInputText\n";
				break;
			}

			bstrCurrentValue.Append(bstrText[i]);
			hRes = InputTextInElement(m_spHtmlElement, bstrCurrentValue);
			if (FAILED(hRes))
			{
				break;
			}

			hRes = m_spPlugin->FireEventOnElement(m_spHtmlElement, CComBSTR("onkeyup"), bstrText[i], vbAsync);
			if (FAILED(hRes))
			{
				traceLog << "Failed to rise onkeyup event in CElement::IeInputText\n";
				break;
			}
		}
	}
	else
	{
		CComQIPtr<IHTMLElement2> spElem2 = m_spHtmlElement;
		if (spElem2 == NULL)
		{
			hRes = E_NOINTERFACE;
		}
		else
		{
			// Set focus on <input file/> control.
			hRes = spElem2->focus();

			if (SUCCEEDED(hRes))
			{
				hRes = InputTextInElement(m_spHtmlElement, bstrText);
			}
		}
	}

	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::IeInputText\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "Failed to rise key events in CElement::IeInputText\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	// Take the focus away from html element.
	BOOL bGenerateOnChange = (bInitValue && (bstrInitialValue != bstrText));
	HRESULT hFocusAway = m_spPlugin->SetFocusAwayFromElement(m_spHtmlElement, bGenerateOnChange, vbAsync);
	if (FAILED(hFocusAway))
	{
		// We can live without focus. Continue.
		traceLog << "Failed to take focus away from element in CElement::IeInputText\n";
	}

	return HRES_OK;
}


HRESULT CElement::InputTextInElement(IHTMLElement* pElement, BSTR bstrText)
{
	ATLASSERT(pElement != NULL);
	ATLASSERT(bstrText != NULL);

	CComQIPtr<IHTMLInputElement> spInputElement;
	HRESULT hRes = pElement->QueryInterface(&spInputElement);

	if (FAILED(hRes))
	{
		if (E_NOINTERFACE != hRes)
		{
			traceLog << "Query for IHTMLInputElement failed in CElement::InputTextInElement with code " << hRes << "\n";
			return hRes;
		}
		else
		{
			// Query for IHTMLTextAreaElement.
			CComQIPtr<IHTMLTextAreaElement> spTextArea;
			hRes = pElement->QueryInterface(&spTextArea);
			if (FAILED(hRes))
			{
				traceLog << "Query for IHTMLTextAreaElement failed in CElement::InputTextInElement with code " << hRes << "\n";
				return hRes;
			}

			hRes = spTextArea->put_value(bstrText);
			if (FAILED(hRes))
			{
				traceLog << "IHTMLTextAreaElement::put_value failed with code " << hRes << " in CElement::InputTextInElement\n";
				return hRes;
			}
		}
	}
	else	// <input> element case.
	{
		// Get the type of the input element.
		CComBSTR bstrType;
		hRes = spInputElement->get_type(&bstrType);

		if (FAILED(hRes) || (bstrType == NULL))
		{
			traceLog << "IHTMLInputElement::get_type failed with code " << hRes << " in CElement::InputTextInElement\n";
			return hRes;
		}

		if (!_wcsicmp(bstrType, L"file"))
		{
			// <input type=file> case. 
			CComQIPtr<IHTMLInputFileElement> spInputFile;
			hRes = spInputElement->QueryInterface(&spInputFile);
			if (FAILED(hRes))
			{
				traceLog << "Query for IHTMLInputFileElement failed in CElement::InputTextInElement with code " << hRes << "\n";
				return hRes;
			}

			CComQIPtr<IAccessible> spAcc = HtmlHelpers::HtmlElementToAccessible(pElement);
			if (spAcc != NULL)
			{
				HWND hIeWnd;
				hRes = ::WindowFromAccessibleObject(spAcc, &hIeWnd);
				if (SUCCEEDED(hRes) && ::IsWindow(hIeWnd))
				{
					ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

					USES_CONVERSION;
					if (!Keyboard::ClearInputFileControl(hIeWnd) ||
						FAILED(m_spPlugin->PostInputText(HandleToLong(hIeWnd), W2T(bstrText), TRUE)))
					{
						traceLog << "Keyboard::InputText failed in CElement::InputTextInElement\n";
						return E_FAIL;
					}
				}
				else
				{
					traceLog << "WindowFromAccessibleObject failed with code " << hRes << " in CElement::InputTextInElement\n";
					return E_FAIL;
				}
			}
			else
			{
				traceLog << "HtmlHelpers::HtmlElementToAccessible failed in CElement::InputTextInElement\n";
				return E_FAIL;
			}
		}
		else
		{
			// <input type=text> or // <input type=password> case.
			ATLASSERT(!_wcsicmp(bstrType, L"password") || !_wcsicmp(bstrType, L"text"));

			hRes = spInputElement->put_value(bstrText);
			if (FAILED(hRes))
			{
				traceLog << "IHTMLInputElement::put_value failed with code " << hRes << " in CElement::InputTextInElement\n";
				return hRes;
			}
		}
	}

	return S_OK;
}


HRESULT CElement::HardwareInputText(BSTR bstrText, BOOL bIsInputFileType)
{
	ATLASSERT(bstrText        != NULL);
	ATLASSERT(m_spPlugin      != NULL);
	ATLASSERT(m_spHtmlElement != NULL);

	// Click the element to focus in.
	ClickCallContext clickCtx(FALSE, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
	BOOL    bClickPosted = FALSE;
	HRESULT hClickRes    = Click(clickCtx, FALSE, &bClickPosted);

	if (HRES_FAIL == hClickRes)
	{
		traceLog << "Click failed in CElement::HardwareInputText with code HRES_FAIL\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
		return hClickRes;
	}
	else if (HRES_BRWS_CONNECTION_LOST_ERR == hClickRes)
	{
		traceLog << "Click failed in CElement::HardwareInputText with code HRES_BRWS_CONNECTION_LOST_ERR\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMET_INPUT_TEXT);
		return hClickRes;
	}
	else
	{
		ATLASSERT(HRES_OK == hClickRes);
	}

	if (bClickPosted)
	{
		// Get the value of the element.
		CComBSTR bstrInitialValue;
		BOOL bInitValue = GetElemValue(m_spHtmlElement, &bstrInitialValue);

		CComQIPtr<IHTMLElement2> spElem2 = m_spHtmlElement;
		if (spElem2 == NULL)
		{
			traceLog << "Can not get IHTMLElement2 in CElement::HardwareInputText\n";
			return E_NOINTERFACE;
		}

		HRESULT hRes = spElem2->focus();
		if (FAILED(hRes))
		{
			traceLog << "IHTMLElement::focus failed in CElement::HardwareInputText with code " << hRes << "\n";
			return hRes;
		}

		// Remove any previous text using IE events.
		hRes = InputTextInElement(m_spHtmlElement, CComBSTR(L""));
		if (FAILED(hRes))
		{
			traceLog << "InputTextInElement failed in CElement::HardwareInputText with code " << hRes << "\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			return hRes;
		}

		CComQIPtr<IAccessible> spAcc = HtmlHelpers::HtmlElementToAccessible(m_spHtmlElement);
		if (spAcc != NULL)
		{
			HWND hIeWnd;
			hRes = ::WindowFromAccessibleObject(spAcc, &hIeWnd);
			if (SUCCEEDED(hRes) && ::IsWindow(hIeWnd))
			{
				ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

				USES_CONVERSION;
				if (FAILED(m_spPlugin->PostInputText(HandleToLong(hIeWnd), W2T(bstrText), bIsInputFileType)))
				{
					traceLog << "PostInputText failed in CElement::HardwareInputText\n";
					return E_FAIL;
				}
			}
			else
			{
				traceLog << "WindowFromAccessibleObject failed with code " << hRes << " in CElement::HardwareInputText\n";
				return E_FAIL;
			}
		}
		else
		{
			traceLog << "HtmlHelpers::HtmlElementToAccessible failed in CElement::HardwareInputText\n";
			return E_FAIL;
		}

		return HRES_OK;
	}
	else
	{
		// Remove any previous text using IE events.
		HRESULT hRes = InputTextInElement(m_spHtmlElement, CComBSTR(L""));
		if (FAILED(hRes))
		{
			traceLog << "InputTextInElement failed in CElement::InputText with code " << hRes << "\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			return hRes;
		}

		ATLASSERT(m_spCore        != NULL);
		ATLASSERT(m_spPlugin      != NULL);
		ATLASSERT(m_spHtmlElement != NULL);

		USES_CONVERSION;
		BOOL bRes = Keyboard::InputText(W2T(bstrText));
		if (!bRes)
		{
			traceLog << "Keyboard::InputText failed in CElement::InputText\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, INPUT_TEXT_METHOD, IDH_ELEMET_INPUT_TEXT);
			return HRES_FAIL;
		}

		return HRES_OK;
	}
}


HRESULT CElement::IsHtmlEditBox(IHTMLElement* pElement, HTML_EDIT_BOX_TYPE* pEditBoxType)
{
	ATLASSERT(pElement     != NULL);
	ATLASSERT(pEditBoxType != NULL);

	*pEditBoxType = HTML_NOT_EDIT_BOX;

	CComBSTR bstrTagName;
	HRESULT hRes = pElement->get_tagName(&bstrTagName);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLElement::get_tagName failed in CElement::IsHtmlEditBox with code " << hRes <<  "\n";
		return hRes;
	}

	if (!_wcsicmp(L"textarea", bstrTagName))
	{
		*pEditBoxType = HTML_MULTILINE_EDIT_BOX;
		return S_OK;
	}

	CComQIPtr<IHTMLInputElement> spInput;
	hRes = pElement->QueryInterface(&spInput);
	if (FAILED(hRes))
	{
		traceLog << "Query for IHTMLInputElement failed in CElement::IsHtmlEditBox with code " << hRes <<  "\n";
		return hRes;
	}

	if (spInput != NULL)
	{
		CComBSTR bstrType;
		HRESULT hRes = spInput->get_type(&bstrType);
		if (S_OK == hRes)
		{
			BOOL bIsEdit = FALSE;
			if (!_wcsicmp(L"text", bstrType))
			{
				*pEditBoxType = HTML_TEXT_EDIT_BOX;
				bIsEdit       = TRUE;
			}
			else if (!_wcsicmp(L"password", bstrType))
			{
				*pEditBoxType = HTML_PASSWORD_EDIT_BOX;
				bIsEdit       = TRUE;
			}
			else if (!_wcsicmp(L"file", bstrType))
			{
				*pEditBoxType = HTML_FILE_EDIT_BOX;
				bIsEdit       = TRUE;
			}

			return bIsEdit ? S_OK : S_FALSE;
		}
		else
		{
			traceLog << "Can not get the type of the input element in CElement::IsHtmlEditBox\n";
			return hRes;
		}
	}

	return S_FALSE;
}


STDMETHODIMP CElement::Select(VARIANT vItems)
{
	FIRE_CANCEL_REQUEST();

	SelectCallContext context(SELECT_METHOD, 0, IDH_ELEMENT_SELECT);
	return Select(vItems, CComVariant(), context);
}


STDMETHODIMP CElement::AddSelection(VARIANT vItems)
{
	FIRE_CANCEL_REQUEST();

	SelectCallContext context(ADD_SELECTION_METHOD, Common::ADD_SELECTION, IDH_ELEMENT_ADD_SELECTION);
	return Select(vItems, CComVariant(), context);
}


STDMETHODIMP CElement::SelectRange(VARIANT vStart, VARIANT vEnd)
{
	FIRE_CANCEL_REQUEST();

	SelectCallContext context(SELECT_RANGE_METHOD, 0, IDH_ELEMENT_SELECT_RANGE);
	return Select(vStart, vEnd, context);
}


STDMETHODIMP CElement::AddSelectionRange(VARIANT vStart, VARIANT vEnd)
{
	FIRE_CANCEL_REQUEST();

	SelectCallContext context(ADD_SELECTION_RANGE_METHOD, Common::ADD_SELECTION, IDH_ELEMENT_ADD_SEL_RANGE);
	return Select(vStart, vEnd, context);
}


HRESULT CElement::Select(const VARIANT& vStartItems, const VARIANT& vEndItems, const SelectCallContext& context)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// Check parameter. If vEndItems is not empty that means range select.
	if ((!Common::IsValidOptionVariant(vStartItems)) ||
	    ((VT_EMPTY != vEndItems.vt) && !Common::IsValidOptionVariant(vEndItems)))
	{
		traceLog << "Invalid parameters in CElement::Select\n";
		SetComErrorMessage(IDS_INVALID_PARAM_LIST_IN_METHOD, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID))
	{
		traceLog << "Invalid state in CElement::Select\n";
		return HRES_FAIL;
	}

	// Get the loadTimeout Core property value.
	LONG    nLoadTimeout = 0;
	HRESULT hRes = m_spCore->get_loadTimeout(&nLoadTimeout);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeout failed for CElement::Select with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the loadTimeoutIsError property.
	VARIANT_BOOL vbLoadTimeoutIsErr;
	hRes = m_spCore->get_loadTimeoutIsError(&vbLoadTimeoutIsErr);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeoutIsError failed for CElement::Select with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the searchTimeout property.
	LONG nSearchTimeout;
	hRes = m_spCore->get_searchTimeout(&nSearchTimeout);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_searchTimeout failed for CElement::Select with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	VARIANT_BOOL vbSelectAsync = VARIANT_FALSE;
	hRes = m_spCore->get_asyncHtmlEvents(&vbSelectAsync);

	if (FAILED(hRes))
	{
		traceLog << "ICore::get_asyncHtmlEvents failed in CElement::Select with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	try
	{
		LONG nFlags = context.m_nSearchFlags;
		if (VARIANT_TRUE == vbSelectAsync)
		{
			nFlags |= Common::PERFORM_ASYNC_ACTION;
		}

		SelectOptionsInContainer selecter(m_spHtmlElement, m_spPlugin, vStartItems, vEndItems, nFlags, m_spCore);
		BOOL bLoadTimeout = FinderInTimeout::Find(&selecter, nSearchTimeout, nLoadTimeout, vbLoadTimeoutIsErr, this, this);

		if (bLoadTimeout)
		{
			SetLastErrorCode(ERR_LOAD_TIMEOUT);
			if (VARIANT_TRUE == vbLoadTimeoutIsErr)
			{
				// The script will throw an exception.
				traceLog << "Page load timeout for " << context.m_szFunctionName << "\n";
				SetComErrorMessage(IDS_SELECT_OPTIONS_LOAD_TIMEOUT, context.m_dwHelpContextID);
				return HRES_LOAD_TIMEOUT_ERR;
			}
		}

		if (!selecter.IsSelectionExecuted())
		{
			traceLog << "m_spPlugin->SelectOptions failed with code HRES_NOT_FOUND_ERR in CElement::Select\n";
			SetComErrorMessage(IDS_OPTION_ELEMENT_NOT_FOUND, context.m_dwHelpContextID);
			SetLastErrorCode(ERR_NOT_FOUND);
			return HRES_NOT_FOUND_ERR;
		}
	}
	catch (const ExceptionServices::OperationNotAllowedException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);
		return HRES_OPERATION_NOT_APPLICABLE;
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}
	catch (const ExceptionServices::BrowserDisconnectedException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}
	catch (const ExceptionServices::IndexOutOfBoundException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_INDEX_OUT_OF_BOUNDS, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_INDEX_OUT_OF_BOUND);
		return HRES_INDEX_OUT_OF_BOUND_ERR;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, context.m_szFunctionName, context.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CElement::ClearSelection(void)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (!IsValidState(IDS_METHOD_CALL_FAILED, CLEAR_SELECTION_METHOD, IDH_ELEMENT_CLEAR_SELECTION))
	{
		traceLog << "Invalid state in CElement::ClearSelection\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->ClearSelection(m_spHtmlElement);
	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::ClearSelection\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_CLEAR_SELECTION);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else if (HRES_OPERATION_NOT_APPLICABLE == hRes)
		{
			traceLog << "Select is not applicable to the current html element in CElement::ClearSelection\n";
			SetComErrorMessage(IDS_ERR_INVALID_OPERATION, IDH_ELEMENT_CLEAR_SELECTION);
			SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);
			return HRES_OPERATION_NOT_APPLICABLE;
		}
		else
		{
			traceLog << "m_spPlugin->SelectOptions failed with code " << hRes << " in CElement::ClearSelection\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, CLEAR_SELECTION_METHOD, IDH_ELEMENT_CLEAR_SELECTION);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


STDMETHODIMP CElement::get_uiName(BSTR* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "Invalid parameters in CElement::get_uiName\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, UI_NAME_PROPERTY, IDH_ELEMENT_UI_NAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, UI_NAME_PROPERTY, IDH_ELEMENT_UI_NAME))
	{
		traceLog << "Invalid state in CElement::get_uiName\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->GetElementText(m_spHtmlElement, pVal);
	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::get_uiName\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_UI_NAME);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spPlugin->GetElementText failed with code " << hRes << " in CElement::get_uiName\n";
			SetComErrorMessage(IDS_PROPERTY_FAILED, UI_NAME_PROPERTY, IDH_ELEMENT_UI_NAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


STDMETHODIMP CElement::Highlight(void)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (!IsValidState(IDS_METHOD_CALL_FAILED, HIGHLIGHT_METHOD, IDH_ELEMENT_HIGHLIGHT))
	{
		traceLog << "Invalid state in CElement::Highlight\n";
		return HRES_FAIL;
	}

	HRESULT hRes = S_OK;
	try
	{
		HWND hIeServerWnd = NULL;
		LONG nWnd    = 0;
		hRes         = m_spPlugin->GetIEServerWnd(m_spHtmlElement, &nWnd);
		hIeServerWnd = static_cast<HWND>(LongToHandle(nWnd));

		if (FAILED(hRes) || !::IsWindow(hIeServerWnd))
		{
			throw CreateException(_T("IExplorerPlugin::GetIEServerWnd failed in CElement::Highlight"));
		}

		// Get the top level IE window.
		ATLASSERT(Common::GetWndClass(hIeServerWnd) == _T("Internet Explorer_Server"));
		HWND hTopWnd = Common::GetTopParentWnd(hIeServerWnd);

		//ATLASSERT(Common::GetWndClass(hTopWnd) == _T("IEFrame"));

		// Put the web browser window in foreground.
		BOOL bRes = PutWindowInForeground(hTopWnd);
		if (!bRes)
		{
			throw CreateException(_T("PutWindowInForeground failed in CElement::Highlight"));
		}

		// Show the html element.
		hRes = m_spHtmlElement->scrollIntoView(CComVariant(TRUE));
		if (FAILED(hRes))
		{
			throw CreateException(_T("IHTMLElement::scrollIntoView failed in CElement::Highlight"));
		}

		// Get the screen coordinates of the html element.
		long nLeft, nRight, nTop, nBottom;
		hRes = m_spPlugin->GetElementScreenRect(m_spHtmlElement, &nLeft, &nTop, &nRight, &nBottom);
		if (FAILED(hRes))
		{
			throw CreateException(_T("IExplorerPlugin::GetElementScreenRect failed in CElement::Highlight"));
		}

		// Highlight the html element.
		const int NUMBER_OF_ITERATIONS     =   7;
		const int DELAY_BETWEEN_ITERATIONS = 125;
		for (int i = 0; i < 2 * NUMBER_OF_ITERATIONS; ++i)
		{
			bRes = DrawRectangleInWindow(hIeServerWnd, nLeft, nRight, nTop, nBottom);
			if (!bRes)
			{
				throw CreateException(_T("DrawRectangleInWindow failed in CElement::Highlight"));
			}

			this->Sleep(DELAY_BETWEEN_ITERATIONS);
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << except << " Error code: " << hRes << "\n";
			traceLog << "Connection with the browser lost while calling in CElement::Highlight\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_HIGHLIGHT);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << except << " Error code: " << hRes << "\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, HIGHLIGHT_METHOD, IDH_ELEMENT_HIGHLIGHT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


BOOL CElement::DrawRectangleInWindow(HWND hWnd, long nLeft, long nRight, long nTop, long nBottom)
{
	ATLASSERT(::IsWindow(hWnd));

	POINT topLeft     = { nLeft,     nTop };
	POINT bottomRight = { nRight, nBottom };
	::ScreenToClient(hWnd, &topLeft);
	::ScreenToClient(hWnd, &bottomRight);

	nLeft   = topLeft.x;
	nTop    = topLeft.y;
	nRight  = bottomRight.x;
	nBottom = bottomRight.y;

	HDC hDC = ::GetWindowDC(hWnd);
	if (NULL == hDC)
	{
		traceLog << "GetWindowDC failed in CElement::DrawRectangleInWindow\n";
		return FALSE;
	}

	// Create a red brush.
	HBRUSH hBrushRed = ::CreateSolidBrush(RGB(0, 0, 0));
	if (NULL == hBrushRed)
	{
		traceLog << "CreateSolidBrush failed in CElement::DrawRectangleInWindow\n";
		::ReleaseDC(hWnd, hDC);
		return FALSE;
	}

	const int HIGHLIGHT_LINE_WIDTH = 3;
	const DWORD DRAW_OPERATION = DSTINVERT; //PATINVERT
	HBRUSH hBrushOld = static_cast<HBRUSH>(::SelectObject(hDC, hBrushRed));

    PatBlt(hDC, nLeft, nTop, nRight - nLeft, HIGHLIGHT_LINE_WIDTH, DRAW_OPERATION);
	PatBlt(hDC, nRight, nBottom - HIGHLIGHT_LINE_WIDTH,
	       -(nRight - nLeft),HIGHLIGHT_LINE_WIDTH, DRAW_OPERATION);
    PatBlt(hDC, nRight - HIGHLIGHT_LINE_WIDTH,	nTop + HIGHLIGHT_LINE_WIDTH,
	       HIGHLIGHT_LINE_WIDTH, nBottom - nTop - 2 * HIGHLIGHT_LINE_WIDTH,	DRAW_OPERATION);
	PatBlt(hDC, nLeft, nBottom - HIGHLIGHT_LINE_WIDTH,
	       HIGHLIGHT_LINE_WIDTH, -(nBottom - nTop - 2 * HIGHLIGHT_LINE_WIDTH), DRAW_OPERATION);

	// Clean up.
	::SelectObject(hDC, hBrushOld);
	::DeleteObject(hBrushRed);
	::ReleaseDC(hWnd, hDC);

	return TRUE;
}


STDMETHODIMP CElement::get_tagName(BSTR* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "Invalid parameters in CElement::get_tagName\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_TAG_NAME_PROPERTY, IDH_ELEMENT_TAGNAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, ELEMENT_TAG_NAME_PROPERTY, IDH_ELEMENT_TAGNAME))
	{
		traceLog << "Invalid state in CElement::get_tagName\n";
		return HRES_FAIL;
	}

	CComBSTR bstrTagName;
	HRESULT  hRes = m_spHtmlElement->get_tagName(&bstrTagName);

	if (FAILED(hRes) || (bstrTagName == NULL))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::get_tagName\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_ELEMENT_TAGNAME);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spHtmlElement->get_tagName failed with code " << hRes << " in CElement::get_tagName\n";
			SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_TAG_NAME_PROPERTY, IDH_ELEMENT_TAGNAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	bstrTagName.ToLower();
	*pVal = bstrTagName.Detach();

	return HRES_OK;
}


HRESULT CElement::GetHandlerAttrText(BSTR bstrAttrName, VARIANT* pVal)
{
	ATLASSERT(bstrAttrName    != NULL);
	ATLASSERT(pVal            != NULL);
	ATLASSERT(m_spHtmlElement != NULL);
	ATLASSERT(VT_EMPTY == pVal->vt);

	try
	{
		CComQIPtr<IHTMLDOMNode> spNode = m_spHtmlElement;
		if (!spNode)
		{
			traceLog << "Cannot get IHTMLDOMNode in CElement::GetHandlreAttrText";
			throw -1;
		}

		CComQIPtr<IDispatch> spListAttrDisp;
		HRESULT              hRes = spNode->get_attributes(&spListAttrDisp);

		CComQIPtr<IHTMLAttributeCollection> spAttributesList = spListAttrDisp;
		if (FAILED(hRes) || !spAttributesList)
		{
			traceLog << "Cannot get spAttributesList in CElement::GetHandlreAttrText";
			throw -1;
		}

		CComVariant          vAttributeName = bstrAttrName;
		CComQIPtr<IDispatch> spAttributeDisp;

		hRes = spAttributesList->item(&vAttributeName, &spAttributeDisp);

		CComQIPtr<IHTMLDOMAttribute> spAttribute = spAttributeDisp;
		if (FAILED(hRes) || !spAttribute)
		{
			traceLog << "Cannot get spAttribute in CElement::GetHandlreAttrText";
			throw -1;
		}

		CComVariant vAttrValue;
		hRes = spAttribute->get_nodeValue(&vAttrValue);

		if (FAILED(hRes))
		{
			traceLog << "Cannot get vAttrValue in CElement::GetHandlreAttrText";
			throw -1;
		}

		if ((vAttrValue.vt != VT_BSTR) || (NULL == vAttrValue.bstrVal))
		{
			vAttrValue = L"";
		}

		hRes = vAttrValue.Detach(pVal);
		if (FAILED(hRes))
		{
			traceLog << "vAttrValue.Detach failed in CElement::GetHandlreAttrText";
			throw -1;
		}

		return HRES_OK;
	}
	catch (int)
	{
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_GET_ATTRIBUTE_METHOD, IDH_ELEMENT_GET_ATTRIBUTE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
}


STDMETHODIMP CElement::GetAttribute(BSTR bstrAttrName, VARIANT* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((NULL == pVal) || (bstrAttrName == NULL))
	{
		traceLog << "Invalid parameters in CElement::GetAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_GET_ATTRIBUTE_METHOD, IDH_ELEMENT_GET_ATTRIBUTE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::GetAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_GET_ATTRIBUTE_METHOD, IDH_ELEMENT_GET_ATTRIBUTE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	HRESULT hRes = m_spHtmlElement->getAttribute(bstrAttrName, 0, pVal); // Case insensitive search of attribute.
	if (SUCCEEDED(hRes) && (VT_DISPATCH == pVal->vt))
	{
		CComQIPtr<IHTMLStyle> spStyle = pVal->pdispVal;
		if (spStyle)
		{
			::VariantClear(pVal);
			
			CComBSTR bstrStyle;
			hRes = spStyle->get_cssText(&bstrStyle);
			if (FAILED(hRes))
			{
				traceLog << "spStyle->toString failed with code=" << hRes <<"\n";
				return hRes;
			}

			if (!bstrStyle)
			{
				bstrStyle = L"";
			}

			CComVariant vStyle = bstrStyle;
			return vStyle.Detach(pVal);
		}

		::VariantClear(pVal); // Release the IDispatch inside pVal variant.
		return GetHandlerAttrText(bstrAttrName, pVal);
	}

	if (FAILED(hRes) ||
	    ((pVal->vt != VT_BOOL) && (pVal->vt != VT_BSTR) && (pVal->vt != VT_NULL) && (pVal->vt != VT_I4)) ||
		((VT_BSTR == pVal->vt) && (NULL == pVal->bstrVal)))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::GetAttribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, IDH_ELEMENT_GET_ATTRIBUTE);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spHtmlElement->getAttribute failed with code " << hRes << " in CElement::GetAttribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_GET_ATTRIBUTE_METHOD, IDH_ELEMENT_GET_ATTRIBUTE);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	// If the attribute is not set return emtpy string.
	if (VT_NULL == pVal->vt)
	{
		pVal->vt      = VT_BSTR;
		pVal->bstrVal = CComBSTR("").Detach();
	}

	return HRES_OK;
}


STDMETHODIMP CElement::SetAttribute(BSTR bstrAttrName, VARIANT varAttrValue)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((bstrAttrName == NULL) ||
		((varAttrValue.vt != VT_BOOL) && (varAttrValue.vt != VT_BSTR) && (varAttrValue.vt != VT_I4)) ||
		((VT_BSTR == varAttrValue.vt) && (NULL == varAttrValue.bstrVal)))
	{
		traceLog << "Invalid parameters in CElement::SetAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_SET_ATTRIBUTE_METHOD, IDH_ELEMENT_SET_ATTRIBUTE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::SetAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_SET_ATTRIBUTE_METHOD, IDH_ELEMENT_SET_ATTRIBUTE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}


	HRESULT hRes = m_spHtmlElement->setAttribute(bstrAttrName, varAttrValue, 0); // Case insensitive search of attribute.
	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::SetAttribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, IDH_ELEMENT_SET_ATTRIBUTE);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spHtmlElement->setAttribute failed with code " << hRes << " in CElement::SetAttribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_SET_ATTRIBUTE_METHOD, IDH_ELEMENT_SET_ATTRIBUTE);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


STDMETHODIMP CElement::RemoveAttribute(BSTR bstrAttrName)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (bstrAttrName == NULL)
	{
		traceLog << "Invalid parameters in CElement::RemoveAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_REMOVE_ATTRIBUTE_METHOD, IDH_ELEMENT_REMOVE_ATTRIBUTE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::RemoveAttribute\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_REMOVE_ATTRIBUTE_METHOD, IDH_ELEMENT_REMOVE_ATTRIBUTE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	VARIANT_BOOL vbSuccess = VARIANT_FALSE;
	HRESULT      hRes      = m_spHtmlElement->removeAttribute(bstrAttrName, 0, &vbSuccess); // Case insensitive search of attribute.

	// vbSuccess is VARIANT_FALSE if the attribute doesn't exist.
	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CElement::RemoveAttribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, IDH_ELEMENT_REMOVE_ATTRIBUTE);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spHtmlElement->get_tagName failed with code " << hRes << " in CElement::SetAtRemoveAttributetribute\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, ELEMENT_REMOVE_ATTRIBUTE_METHOD, IDH_ELEMENT_REMOVE_ATTRIBUTE);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


STDMETHODIMP CElement::FindParentElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	ATLASSERT(m_spHtmlElement != NULL);
	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CElement::FindParentElement\n";
		SetComErrorMessage(IDS_ERR_FIND_PARENT_ELEMENT_INVALID_PARAM, IDH_ELEMENT_FIND_PARENT);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_PARENT_ELEM;
	SearchContext sc("CElement::FindParentElement", nSearchFlags, IDH_ELEMENT_FIND_PARENT,
					 IDS_ERR_FIND_PARENT_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_PARENT_ELEMENT_FAILED,
					 IDS_FIND_PARENT_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CElement::FindParentElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_FIND_PARENT);
			SetLastErrorCode(ERR_FAIL);

			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::get_isChecked(VARIANT_BOOL* pIsChecked)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pIsChecked)
	{
		traceLog << "pVal is NULL in CElement::get_isChecked\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_ISCHECKED_PROPERTY, IDH_ELEMENT_IS_CHECKED);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::get_isChecked\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_ISCHECKED_PROPERTY, IDH_ELEMENT_IS_CHECKED);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	if (!IsCheckable(m_spHtmlElement))
	{
		traceLog << "get_isChecked is not applicable to the current html element in CElement::get_isChecked\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, IDH_ELEMENT_IS_CHECKED);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);

		return HRES_OPERATION_NOT_APPLICABLE;
	}

	CComQIPtr<IHTMLInputElement> spInputElement = m_spHtmlElement;
	ATLASSERT(spInputElement != NULL); // Already checked by IsCheckable.

	VARIANT_BOOL vbChecked;
	HRESULT      hRes = spInputElement->get_checked(&vbChecked);

	if (FAILED(hRes))
	{
		traceLog << "IHTMLInputElement::get_checked failed with code " << hRes << " in CElement::get_isChecked\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_ISCHECKED_PROPERTY, IDH_ELEMENT_IS_CHECKED);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	*pIsChecked = vbChecked;
	return HRES_OK;
}


BOOL CElement::IsCheckable(CComQIPtr<IHTMLElement> spElement, BOOL bRadioIsNotCheckable)
{
	ATLASSERT(spElement != NULL);

	CComQIPtr<IHTMLInputElement> spInputElement = spElement;
	if (spInputElement == NULL)
	{
		traceLog << "spInputElement is NULL in CElement::IsCheckable\n";
		return FALSE;
	}

	// Get the type of the input element.
	CComBSTR bstrType;
	HRESULT  hRes = spInputElement->get_type(&bstrType);

	if (FAILED(hRes) || (bstrType == NULL))
	{
		traceLog << "IHTMLInputElement::get_type failed with code " << hRes << " in CElement::IsCheckable\n";
		return FALSE;
	}

	if (!_wcsicmp(bstrType, L"checkbox") || (!bRadioIsNotCheckable && !_wcsicmp(bstrType, L"radio")))
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


STDMETHODIMP CElement::Check()
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::Check\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_CHECK_METHOD, IDH_ELEMENT_CHECK);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	if (!IsCheckable(m_spHtmlElement))
	{
		traceLog << "get_isChecked is not applicable to the current html element in CElement::Check\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, IDH_ELEMENT_CHECK);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);

		return HRES_OPERATION_NOT_APPLICABLE;
	}

	CComQIPtr<IHTMLInputElement> spInputElement = m_spHtmlElement;
	ATLASSERT(spInputElement != NULL); // Already checked by IsCheckable.

	VARIANT_BOOL vbChecked;
	HRESULT      hRes = spInputElement->get_checked(&vbChecked);

	if (FAILED(hRes))
	{
		traceLog << "IHTMLInputElement::get_checked failed with code " << hRes << " in CElement::Check\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_CHECK_METHOD, IDH_ELEMENT_CHECK);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	hRes = HRES_OK;

	if (vbChecked != VARIANT_TRUE)
	{
		// The element is not checked. Check it by click.
		ClickCallContext clickCtx(FALSE, ELEMENT_CHECK_METHOD, IDH_ELEMENT_CHECK);
		hRes = Click(clickCtx, TRUE);
	}

	return hRes;
}


STDMETHODIMP CElement::Uncheck()
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (m_spHtmlElement == NULL)
	{
		traceLog << "m_spHtmlElement is NULL in CElement::Uncheck\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_UNCHECK_METHOD, IDH_ELEMENT_UNCHECK);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	// Radio elements are not un-checkable. The get unchecked when other radio is pressed.
	if (!IsCheckable(m_spHtmlElement, TRUE))
	{
		traceLog << "get_isChecked is not applicable to the current html element in CElement::Uncheck\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, IDH_ELEMENT_UNCHECK);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);

		return HRES_OPERATION_NOT_APPLICABLE;
	}

	CComQIPtr<IHTMLInputElement> spInputElement = m_spHtmlElement;
	ATLASSERT(spInputElement != NULL); // Already checked by IsCheckable.

	VARIANT_BOOL vbChecked;
	HRESULT      hRes = spInputElement->get_checked(&vbChecked);

	if (FAILED(hRes))
	{
		traceLog << "IHTMLInputElement::get_checked failed with code " << hRes << " in CElement::Uncheck\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, ELEMENT_UNCHECK_METHOD, IDH_ELEMENT_UNCHECK);
		SetLastErrorCode(ERR_FAIL);

		return HRES_FAIL;
	}

	hRes = HRES_OK;

	if (VARIANT_TRUE == vbChecked)
	{
		// The element is checked. Uncheck it by click.
		ClickCallContext clickCtx(FALSE, ELEMENT_UNCHECK_METHOD, IDH_ELEMENT_UNCHECK);
		hRes = Click(clickCtx, TRUE);
	}

	return hRes;
}


STDMETHODIMP CElement::get_selectedOption(IElement** ppSelectedOption)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppSelectedOption)
	{
		traceLog << "ppSelectedOption is null in CElement::get_selectedOption\n";
		SetComErrorMessage(IDS_ERR_SELECTED_OPTION_INVALID_PARAM, IDH_ELEMENT_SELECTED_OPTION);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_SELECTED_OPTION;
	SearchContext sc("CElement::get_selectedOption", nSearchFlags, IDH_ELEMENT_SELECTED_OPTION,
					 IDS_ERR_SELECTED_OPTION_INVALID_PARAM, IDS_ERR_SELECTED_OPTION_FAILED,
					 IDS_SELECTED_OPTION_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, NULL, NULL, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppSelectedOption);
		if (NULL == *ppSelectedOption)
		{
			traceLog << "Can not create an Element object in CElement::get_selectedOption\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_ELEMENT_SELECTED_OPTION);
			SetLastErrorCode(ERR_FAIL);
			hRes = HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CElement::GetAllSelectedOptions(IElementList** ppSelectedOptionsList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppSelectedOptionsList)
	{
		traceLog << "ppSelectedOptionsList is null in CElement::get_selectedOption\n";
		SetComErrorMessage(IDS_ERR_GET_ALL_SEL_OPTIONS_INVALID_PARAM, IDH_ELEMENT_ALL_SELECTED_OPTIONS);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_SELECTED_OPTION | SEARCH_COLLECTION;
	SearchContext sc("CElement::GetAllSelectedOptions", nSearchFlags, IDH_ELEMENT_ALL_SELECTED_OPTIONS,
					 IDS_ERR_GET_ALL_SEL_OPTIONS_INVALID_PARAM, IDS_ERR_GET_ALL_SEL_OPTIONS_FAILED,
					 IDS_GET_ALL_SEL_OPTIONS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(m_spHtmlElement, sc, NULL, NULL, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateList = CreateElemListObject(spResult, ppSelectedOptionsList);
		if (NULL == *ppSelectedOptionsList)
		{
			traceLog << "Can not create an ElementList object in CElement::GetAllSelectedOptions\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_ELEMENT_ALL_SELECTED_OPTIONS);
			SetLastErrorCode(ERR_FAIL);
			hRes = HRES_FAIL;
		}
	}

	return hRes;
}

BOOL CElement::IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID)
{
	if (!BaseLibObject::IsValidState(nPatternID, szMethodName, dwHelpID))
	{
		traceLog << "IsValidState CElement::IsValidState\n";
		return FALSE;
	}
	else
	{
		if (m_spHtmlElement == NULL)
		{
			traceLog << "m_spHtmlElement is NULL in CElement::IsValidState\n";
			SetComErrorMessage(nPatternID, szMethodName, dwHelpID);
			SetLastErrorCode(ERR_FAIL);
			return FALSE;
		}
	}

	return TRUE;
}
