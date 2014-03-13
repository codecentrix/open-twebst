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

#include "StdAfx.h"
#include "resource.h"
#include "DebugServices.h"
#include "CodeErrors.h"
#include "HtmlHelpers.h"
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "FindInTimeout.h"


#undef FIRE_CANCEL_REQUEST
#define FIRE_CANCEL_REQUEST(pThisFinder) \
{\
	CCore* pCore =  static_cast<CCore*>(pThisFinder->m_spCore.p);\
	if (pCore != NULL)\
	{\
		pCore->Fire_CancelRequest();\
		if (pCore->IsCancelPending())\
		{\
			throw CreateExecutionCanceledException(_T("Execution canceled in FinderInTimeout"));\
		}\
	}\
}


// Returns TRUE if loadTimeout expired.
BOOL FinderInTimeout::Find(FinderInTimeout*  pFinder,
                           ULONG             nSearchTimeout,
                           ULONG             nLoadTimeout,
                           VARIANT_BOOL      vbLoadTimeoutIsErr,
						   IAutoClosePopups* pAutoClosePopups,
						   ISleeper*         pSleeper)
{
	// Sleep for a while.
	if (pSleeper != NULL)
	{
		BOOL bDispatchMsg = (nLoadTimeout || nSearchTimeout);
		pSleeper->Sleep(Common::INTERNAL_GLOBAL_PAUSE, bDispatchMsg);
	}

	ATLASSERT(pFinder != NULL);

	DWORD dwLoadingTime = 0;
	DWORD dwSearchTime  = 0;
	BOOL  bLoadTimeout  = FALSE;

	while (TRUE)
	{
		if (!bLoadTimeout && (nLoadTimeout != 0))
		{
			// Wait the page to load.
			DWORD dwWaitToLoadStartTime = ::GetTickCount();
			while (pFinder->IsBrowserLoading())
			{
				FIRE_CANCEL_REQUEST(pFinder);

				if (pAutoClosePopups != NULL)
				{
					pAutoClosePopups->CloseBrowserPopups();
				}

				DWORD dwCurrentTime   = ::GetTickCount();
				DWORD dwElapsedTime   = dwCurrentTime - dwWaitToLoadStartTime;
				dwWaitToLoadStartTime = dwCurrentTime;
				dwLoadingTime += dwElapsedTime;

				if (dwLoadingTime > (nLoadTimeout * TIME_SCALE))
				{
					bLoadTimeout = TRUE;
					break;
				}

				// Sleep for a while.
				if (pSleeper != NULL)
				{
					pSleeper->Sleep(Common::INTERNAL_GLOBAL_PAUSE, TRUE);
				}
			}

			if ((VARIANT_TRUE == vbLoadTimeoutIsErr) && (TRUE == bLoadTimeout))
			{
				// Load timeout.
				return TRUE;
			}
		}

		// Search the element.
		DWORD dwSearchStartTime = ::GetTickCount();
		while (TRUE)
		{
			FIRE_CANCEL_REQUEST(pFinder);

			if (pAutoClosePopups != NULL)
			{
				pAutoClosePopups->CloseBrowserPopups();
			}

			if ((nLoadTimeout != 0) && !bLoadTimeout && pFinder->IsBrowserLoading())
			{
				// Exit and wait again.
				break;
			}

			// Search elements.
			BOOL bFound = pFinder->Find();
			if (bFound)
			{
				// Element(s) found.
				if ((nLoadTimeout != 0) && !bLoadTimeout)
				{
					// Wait the element to be completely loaded.
					ATLASSERT(dwLoadingTime <= (nLoadTimeout * TIME_SCALE));

					DWORD dwLoadTimeoutOnObject = (nLoadTimeout * TIME_SCALE) - dwLoadingTime;
					BOOL  bElemLoadOK           = pFinder->WaitToLoad(dwLoadTimeoutOnObject, pSleeper);

					if (bElemLoadOK)
					{
						return FALSE; // No timeout.
					}
					else
					{
						if (VARIANT_TRUE == vbLoadTimeoutIsErr)
						{
							return TRUE; // Load timeout on element.
						}
						else
						{
							return FALSE;
						}
					}
				}

				return FALSE;
			}

			DWORD dwCurrentTime = ::GetTickCount();
			DWORD dwElapsedTime = dwCurrentTime - dwSearchStartTime;
			dwSearchStartTime   = dwCurrentTime;
			dwSearchTime += dwElapsedTime;

			if (dwSearchTime > nSearchTimeout * TIME_SCALE)
			{
				// Search timeout expired. Return FALSE because TRUE is for load timeout.
				return FALSE;
			}
		}
	}

	return FALSE;
}


BOOL FinderInTimeout::IsBrowserLoading()
{
	ATLASSERT(m_spPlugin != NULL);

	// Get the loading state of the browser.
	VARIANT_BOOL vbIsLoading;
	HRESULT hRes = m_spPlugin->IsLoading(&vbIsLoading);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		throw CreateBrowserDisconnectedException(_T("Connectin with the browser lost in IsBrowserLoading"));
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugIn->IsLoading failed in IsBrowserLoading\n";
		return TRUE; // Assume is still loading.
	}

	CComQIPtr<IHTMLWindow2> spWindow = m_spContainer;
	if (spWindow != NULL)
	{
		// In case of frames wait the frame to be completely loaded too.
		// This is for HTML dialogs.
		// For performance reasons maybe HtmlHelpers::IsWindowReady should be made in IE process thru BHO call.
		try
		{
			if (!HtmlHelpers::IsWindowReady(spWindow))
			{
				// Still loading.
				return TRUE;
			}
		}
		catch (const ExceptionServices::Exception& except)
		{
			traceLog << except << "\n";
			traceLog << "HtmlHelpers::IsWindowReady failed in FinderInTimeout::IsBrowserLoading\n";
			return TRUE; // Assume the frame is still loading.
		}
	}

	return VARIANT_TRUE == vbIsLoading;
}


//////////////////////////////////////////////////////////////////////////////////////////////
BOOL FindObjectInContainer::Find()
{
	m_spResult = NULL;

	ATLASSERT(m_spPlugin    != NULL);
	ATLASSERT((m_bstrTag != NULL) ||
	          ((m_bstrTag == NULL) &&
			  ((m_nSearchFlags & Common::SEARCH_FRAME)  || (m_nSearchFlags & SEARCH_HTML_DIALOG) ||
			   (m_nSearchFlags & SEARCH_MODAL_HTML_DLG) || (m_nSearchFlags & SEARCH_SELECTED_OPTION))));

	HRESULT hRes = S_OK;

	if (m_nSearchFlags & SEARCH_SELECTED_OPTION)
	{
		CComQIPtr<IHTMLElement> spSelect = m_spContainer;

		if (spSelect != NULL)
		{
			if  (m_nSearchFlags & SEARCH_COLLECTION)
			{
				hRes = m_spPlugin->FindAllSelectedOptions(spSelect, &m_spResult);
			}
			else
			{
				CComQIPtr<IHTMLElement> spSelectedOption;

				hRes       = m_spPlugin->FindSelectedOption(spSelect, &spSelectedOption);
				m_spResult = spSelectedOption;
			}
		}
		else
		{
			hRes = E_INVALIDARG;
		}
	}
	else
	{
		// Search the object (element/frame/collection).
		ATLASSERT(m_pVarArgs != NULL);
		hRes = m_spPlugin->FindInContainer(m_spContainer, m_nSearchFlags, m_bstrTag, m_pVarArgs, &m_spResult);
	}

	if (HRES_OPERATION_NOT_APPLICABLE == hRes)
	{
		throw CreateOperationNotAllowedException(_T("Operation is not applicable to the current html element in FindObjectInContainer::Find"));
	}

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		throw CreateBrowserDisconnectedException(_T("Connectin with the browser lost in FindObjectInContainer::Find while calling m_spPlugin->FindInContainer"));
	}

	if (E_INVALIDARG == hRes)
	{
		throw CreateInvalidParamException(_T("m_spPlugin->FindInContainer/FindSelectedOption returned E_INVALIDARG in FindObjectInContainer::Find"));
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->FindInContainer/FindSelectedOption failed in FindObjectInContainer::Find\n";

		// Continue searching in timeout.
		return FALSE;
	}

	if (m_spResult != NULL)
	{
		if (m_nSearchFlags & SEARCH_COLLECTION)
		{
			if (S_FALSE == hRes)
			{
				return FALSE; // Empty collection.
			}
			else
			{
				return TRUE;
			}
		}
		else
		{
			// The element was found.
			return TRUE;
		}
	}
	else
	{
		// The element was not found. Continue searching in timeout.
		ATLASSERT(S_FALSE == hRes);
		return FALSE;
	}
}


BOOL FindObjectInContainer::WaitToLoad(DWORD nLoadTimeout, ISleeper* pSleeper)
{
	if (nLoadTimeout == 0)
	{
		return TRUE;
	}

	if ((m_nSearchFlags & Common::SEARCH_HTML_DIALOG) ||
		(m_nSearchFlags & SEARCH_MODAL_HTML_DLG))
	{
		// Wait only HTML dialogs to be completely loaded. For other types of objects
		// there's no such need because we wait the whole browser to complete.

		CComQIPtr<IHTMLWindow2> spWindow = m_spResult;
		if (spWindow == NULL)
		{
			throw CreateException(_T("spWindow is NULL in FindObjectInContainer::WaitToLoad"));
		}

		DWORD dwWaitToLoadStartTime = ::GetTickCount();

		while (TRUE)
		{
			try
			{
				// For performance reasons maybe HtmlHelpers::IsWindowReady should be made in IE process thru BHO call.
				if (HtmlHelpers::IsWindowReady(spWindow))
				{
					// Window completely loaded.
					return TRUE;
				}
			}
			catch (const ExceptionServices::Exception& except)
			{
				traceLog << except << "\n";
				traceLog << "HtmlHelpers::IsWindowReady failed in FindObjectInContainer::WaitToLoad\n";
			}

			FIRE_CANCEL_REQUEST(this);

			DWORD dwCurrentTime   = ::GetTickCount();
			DWORD dwElapsedTime   = dwCurrentTime - dwWaitToLoadStartTime;

			if (dwElapsedTime > nLoadTimeout)
			{
				// Load timeout.
				return FALSE;
			}

			// Sleep for a while.
			if (pSleeper != NULL)
			{
				pSleeper->Sleep(Common::INTERNAL_GLOBAL_PAUSE, TRUE);
			}
		}
	}
	else
	{
		return TRUE;
	}
}


//////////////////////////////////////////////////////////////////////////////////////////////
BOOL FindFramesInContainer::Find()
{
	ATLASSERT(m_spPlugin    != NULL);
	ATLASSERT(m_ppFrames    == NULL);

	// Search the object (element/frame/collection).
	ATLASSERT(m_pVarArgs != NULL);

	HRESULT hRes = m_spPlugin->FindFramesInContainer(m_spContainer, m_nSearchFlags, m_pVarArgs, &m_nCollectionSize, &m_ppFrames);

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		throw CreateBrowserDisconnectedException(_T("Connectin with the browser lost in FindFramesInContainer::Find while calling m_spPlugin->FindFramesInContainer"));
	}

	if (E_INVALIDARG == hRes)
	{
		throw CreateException(_T("m_spPlugin->FindFramesInContainer returned E_INVALIDARG in FindFramesInContainer::Find"));
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->FindFramesInContainer failed in FindFramesInContainer::Find\n";

		// Continue searching in timeout.
		return FALSE;
	}

	if (m_ppFrames != NULL)
	{
		// The element was found.
		return TRUE;
	}
	else
	{
		// The element was not found. Continue searching in timeout.
		ATLASSERT(S_FALSE == hRes);
		return FALSE;
	}
}


void FindFramesInContainer::Detach(IHTMLWindow2*** pppFrames, LONG* pSize)
{
	ATLASSERT((pppFrames != NULL) && (NULL == *pppFrames));
	ATLASSERT(pSize != NULL);

	*pppFrames = m_ppFrames;
	*pSize     = m_nCollectionSize;

	m_ppFrames        = NULL;
	m_nCollectionSize = 0;
}


FindFramesInContainer::~FindFramesInContainer()
{
	if (m_ppFrames != NULL)
	{
		ATLASSERT(m_nCollectionSize != 0);
		for (int i = 0; i < m_nCollectionSize; ++i)
		{
			m_ppFrames[i]->Release();
			m_ppFrames[i] = NULL;
		}

		::CoTaskMemFree(m_ppFrames);
		m_ppFrames         = NULL;
		m_nCollectionSize  = 0;
	}
}


//////////////////////////////////////////////////////////////////////////////////////////////
BOOL SelectOptionsInContainer::Find()
{
	CComQIPtr<IHTMLElement> spHtmlElement = m_spContainer;

	if (spHtmlElement == NULL)
	{
		throw CreateInvalidParamException(_T("m_spContainer is not a HTML element in SelectOptionsInContainer::Find"));
	}

	HRESULT hRes = m_spPlugin->SelectOptions(spHtmlElement, m_vStartItems, m_vEndItems, m_nSearchFlags);

	if (HRES_OPERATION_NOT_APPLICABLE == hRes)
	{
		throw CreateOperationNotAllowedException(_T("Operation is not applicable to the current html element in SelectOptionsInContainer::Find"));
	}

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		throw CreateBrowserDisconnectedException(_T("Connectin with the browser lost in FindObjectInContainer::Find while calling m_spPlugin->SelectOptions"));
	}

	if (E_INVALIDARG == hRes)
	{
		throw CreateInvalidParamException(_T("m_spPlugin->SelectOptions returned E_INVALIDARG in SelectOptionsInContainer::Find"));
	}

	if (HRES_INDEX_OUT_OF_BOUND_ERR == hRes)
	{
		throw CreateIndexOutOfBoundException(_T("m_spPlugin->SelectOptions returned HRES_INDEX_OUT_OF_BOUND_ERR in SelectOptionsInContainer::Find"));
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->SelectOptions failed in SelectOptionsInContainer::Find\n";

		// Continue searching in timeout.
		return FALSE;
	}
	else
	{
		m_bExecuted = TRUE;
		return TRUE;
	}
}
