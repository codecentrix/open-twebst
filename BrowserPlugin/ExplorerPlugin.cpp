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
#include "Common.h"
#include "ExplorerPlugin.h"
#include "FindElement.h"
#include "HtmlHelpers.h"
#include "FindFrame.h"
#include "LocalElementCollection.h"
#include "Keyboard.h"
#include "ClickAsyncAction.h"
#include "SetFocusAsyncAction.h"
#include "SetFocusAwayAsyncAction.h"
#include "FireEventAsyncAction.h"
#include "SelectAsyncAction.h"

using namespace Common;

#ifdef min
	#undef min
#endif



HRESULT CExplorerPlugin::SetSite(IUnknown* pSite)
{
	if (pSite != NULL)
	{
		return Init(pSite);
	}
	else
	{
		return UnInit();
	}
}


HRESULT CExplorerPlugin::Init(IUnknown* pSite)
{
	traceLog << "CExplorerPlugin::Init begins\n";

	if (!pSite)
	{
		traceLog << "CExplorerPlugin::Init pSite is NULL\n";
		return E_UNEXPECTED;
	}

	HRESULT hRes = IObjectWithSiteImpl<CExplorerPlugin>::SetSite(pSite);
	if (FAILED(hRes))
	{
		traceLog << "IObjectWithSiteImpl::SetSite failed in CExplorerPlugin::Init\n";
		return hRes;
	}

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);

	if (spBrws == NULL)
	{
		traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::Init\n";
		return E_FAIL;
	}

	// On IE7 each tab window belongs to a different thread.
	// On IE6 each top level window belongs to a different thread.
	ATLASSERT((Common::GetWndClass(m_hIeWnd) == _T("Internet Explorer_Server")) ||
	          (::GetCurrentThreadId() == ::GetWindowThreadProcessId(m_hIeWnd, NULL)));

	String sPluginMarshalName;
	BOOL   bName = NewMarshalService::GetMarshalWndName(m_hIeWnd, sPluginMarshalName);

	if (!bName)
	{
		traceLog << "NewMarshalService::GetMarshalWndName failed in CExplorerPlugin::Init\n";
		return E_FAIL;
	}

	// Create accessible window.
	RECT rect = { 0 };
	HWND hWnd = Create(NULL, rect, sPluginMarshalName.c_str(), WS_POPUP);
	if (NULL == hWnd)
	{
		traceLog << "Can NOT create communication window in CExplorerPlugin::Init\n";
		return E_FAIL;
	}

	// Register to get browser events.
	ATLASSERT(0 == m_dwEventsCookie);

	hRes = AtlAdvise(spBrws, GetUnknown(), DIID_DWebBrowserEvents2, &m_dwEventsCookie);
	if (FAILED(hRes))
	{
		traceLog << "Failed to register for browser events in CExplorerPlugin::Init, code=" << hRes << "\n";
		return hRes;
	}

	traceLog << "CExplorerPlugin::Init end\n";
	return S_OK;
}


HRESULT CExplorerPlugin::UnInit()
{
	traceLog << "CExplorerPlugin::UnInit begins\n";

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::UnInit\n";
		return S_FALSE; // Maybe already disposed.
	}

	// Clean up the queue of async actions (if any).
	CleanAyncActionsQueue();

	// Destroy the marshal communication window.
	BOOL bRes = DestroyWindow();
	m_hWnd = NULL;

	if (!bRes)
	{
		traceLog << "DestroyWindow failed in CExplorerPlugin::UnInit\n";
	}

	traceLog << "CExplorerPlugin::UnInit ends\n";
	return IObjectWithSiteImpl<CExplorerPlugin>::SetSite(NULL);
}


STDMETHODIMP CExplorerPlugin::GetNativeBrowser(IWebBrowser2** ppWebBrowser)
{
	traceLog << "CExplorerPlugin::GetNativeBrowser begins\n";
	if (NULL == ppWebBrowser)
	{
		traceLog << "ppWebBrowser is null in CExplorerPlugin::GetNativeBrowser\n";
		return E_INVALIDARG;
	}

	// Get the IWebBrowser2 pointer.
	IWebBrowser2* pWebBrws = NULL;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&pWebBrws);
	if (pWebBrws == NULL)
	{
		traceLog << "Can NOT obtain IWebBrowser2 browser in CExplorerPlugin::GetNativeBrowser\n";
		return hRes;
	}

	*ppWebBrowser = pWebBrws;
	traceLog << "CExplorerPlugin::GetNativeBrowser ends\n";
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::CheckBrowserDescriptor(LONG nSearchFlags, SAFEARRAY* psa)
{
	if (NULL == psa)
	{
		traceLog << "psa is null in CExplorerPlugin::CheckBrowserDescriptor\n";
		return E_INVALIDARG;
	}

	UINT nDimNumber = ::SafeArrayGetDim(psa);

	if (nDimNumber != 1)
	{
		traceLog << "Wrong number of dimensions of safearray in CExplorerPlugin::CheckBrowserDescriptor\n";
		return E_INVALIDARG;
	}

	std::list<DescriptorToken> tokens;
	BOOL bRes = Common::GetDescriptorTokensList(psa, &tokens);

	if (!bRes)
	{
		traceLog << "Can NOT get the list of descriptor tokens in CExplorerPlugin::CheckBrowserDescriptor\n";
		return E_INVALIDARG;
	}

	for (list<DescriptorToken>::const_iterator it = tokens.begin();
	     it != tokens.end(); ++it)
	{
		String sAttributeValue;

		try
		{
			bRes = GetBrowserAttributeValue(it->m_sAttributeName, &sAttributeValue);
			if (FALSE == bRes)
			{
				traceLog << "Can NOT get the value of the attribute " << it->m_sAttributeName << "\n";
				return S_FALSE;
			}
		}
		catch (const ExceptionServices::Exception& except)
		{
			traceLog << except << "\n";
			return S_FALSE;
		}


		// Ignore last slash in URL.
		String sAttributePattern = it->m_sAttributePattern;
		if (!_tcsicmp(it->m_sAttributeName.c_str(), BROWSER_URL_ATTR_NAME))
		{
			StripLastSlash(&sAttributePattern);
			StripLastSlash(&sAttributeValue);
		}

		BOOL bMatch = Common::MatchWildcardPattern(sAttributePattern, sAttributeValue);
		if (DescriptorToken::TOKEN_MATCH == it->m_operator)
		{
			if (!bMatch)
			{
				return S_FALSE;
			}
		}
		else
		{
			ATLASSERT(DescriptorToken::TOKEN_NOT_MATCH == it->m_operator);
			if (bMatch)
			{
				return S_FALSE;
			}
		}
	}

	return S_OK;
}


BOOL CExplorerPlugin::GetBrowserAttributeValue(const String& sAttributeName, String* pAttributeValue)
{
	ATLASSERT(pAttributeValue != NULL);

	if (!_tcsicmp(sAttributeName.c_str(), BROWSER_TITLE_ATTR_NAME))
	{
		return GetBrowserTitle(pAttributeValue);
	}
	else if (!_tcsicmp(sAttributeName.c_str(), BROWSER_URL_ATTR_NAME))
	{
		return GetBrowserUrl(pAttributeValue);
	}
	else if (!_tcsicmp(sAttributeName.c_str(), BROWSER_PID_ATTR_NAME))
	{
		GetBrowserPid(pAttributeValue);
		return TRUE;
	}
	else if (!_tcsicmp(sAttributeName.c_str(), BROWSER_TID_ATTR_NAME))
	{
		GetBrowserTid(pAttributeValue);
		return TRUE;
	}
	else
	{
		traceLog << sAttributeName << " invalid attribute name\n";
		return FALSE;
	}
}


// Returns TRUE if the title of the HTML document was correctly retrieved.
// Returns FALSE if the browser doesn't display a HTML document.
// Could throw an ExceptionServices::Exception in the case of a failure.
BOOL CExplorerPlugin::GetBrowserTitle(String* pTitle)
{
	ATLASSERT(pTitle != NULL);

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		throw CreateException(_T("Can NOT obtain IWebBrowser2 in CExplorerPlugin::GetBrowserTitle"));
	}

	// Get the document.
	CComQIPtr<IDispatch> spDisp;
	hRes = spBrws->get_Document(&spDisp);

	CComQIPtr<IHTMLDocument2> spDocument = spDisp;
	if (spDocument == NULL)
	{
		throw CreateException(_T("Can NOT obtain the document in CExplorerPlugin::GetBrowserTitle"));
	}

	// Get the title.
	CComBSTR bstrTitle;
	hRes = spDocument->get_title(&bstrTitle);
	if (FAILED(hRes) || (bstrTitle == NULL))
	{
		throw CreateException(_T("Can NOT get the title in CExplorerPlugin::GetBrowserTitle"));
	}

	ATLASSERT(bstrTitle != NULL);
	USES_CONVERSION;
	*pTitle = W2T(bstrTitle);
	return TRUE;
}


// Returns TRUE if the URL of the HTML document was correctly retrieved.
// Returns FALSE if the browser doesn't display a HTML document.
// Could throw an ExceptionServices::Exception in the case of a failure.
BOOL CExplorerPlugin::GetBrowserUrl(String* pUrl)
{
	ATLASSERT(pUrl != NULL);

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		throw CreateException(_T("Can NOT obtain IWebBrowser2 in CExplorerPlugin::GetBrowserUrl"));
	}

	CComBSTR bstrUrl;
	hRes = spBrws->get_LocationURL(&bstrUrl);
	if ((FAILED(hRes)))
	{
		traceLog << "IWebBrowser2::get_LocationURL failed in CExplorerPlugin::GetBrowserUrl with code:" << hRes << "\n";
		throw CreateException(_T("IWebBrowser2::get_LocationURL failed in CExplorerPlugin::GetBrowserUrl"));
	}

	ATLASSERT(bstrUrl != NULL);
	USES_CONVERSION;
	*pUrl = W2T(bstrUrl);
	return TRUE;
}


// Could throw an ExceptionServices::Exception if the PID of the IE process can not be correctly retrieved.
void CExplorerPlugin::GetBrowserPid(String* pPid)
{
	ATLASSERT(pPid != NULL);
	DWORD dwCurrentPid = ::GetCurrentProcessId();

	Common::Ostringstream outputStream;
	outputStream << dwCurrentPid;
	*pPid = outputStream.str();
}


// Could throw an ExceptionServices::Exception if the TID of the IE process can not be correctly retrieved.
void CExplorerPlugin::GetBrowserTid(String* pTid)
{
	ATLASSERT(pTid != NULL);
	DWORD dwCurrentTid = ::GetCurrentThreadId();

	Common::Ostringstream outputStream;
	outputStream << dwCurrentTid;
	*pTid =  outputStream.str();
}


STDMETHODIMP CExplorerPlugin::IsLoading(VARIANT_BOOL* pLoading)
{
	if (NULL == pLoading)
	{
		return E_INVALIDARG;
	}

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		return hRes;
	}

	BOOL bIsBusy = FALSE;

	// Get the browser busy state.
	VARIANT_BOOL vbBusy = VARIANT_FALSE;

	// Try get_Busy method first. If it says is not busy then it is so.
	// If it says is busy that might be because of a modal dialog box; in this case check the number of pending downloads.
	hRes = spBrws->get_Busy(&vbBusy);
	if (FAILED(hRes) || (VARIANT_TRUE == vbBusy))
	{
		ATLASSERT(m_nActiveDownloads >= 0);
		if (m_nActiveDownloads > 0)
		{
			// The browser is loading.
			*pLoading = VARIANT_TRUE;

			return S_OK;
		}
	}

	try
	{
		*pLoading = (HtmlHelpers::IsBrowserReady(spBrws) ? VARIANT_FALSE : VARIANT_TRUE);
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		traceLog << "HtmlHelpers::IsBrowserReady failed in CExplorerPlugin::IsLoading\n";
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetTopFrame(IHTMLWindow2** pTopFrame)
{
	if (NULL == pTopFrame)
	{
		traceLog << "pTopFrame is NULL in CExplorerPlugin::GetTopFrame\n";
		return E_INVALIDARG;
	}

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		traceLog << "Can not get the IWebBrowser2 in CExplorerPlugin::GetTopFrame\n";
		return E_FAIL;
	}

	CComQIPtr<IHTMLWindow2> spTopWindow = HtmlHelpers::HtmlWebBrowserToHtmlWindow(spBrws);
	*pTopFrame = spTopWindow.Detach();
	return S_OK;
}


// Return S_FALSE if the element/collection was not found.
STDMETHODIMP CExplorerPlugin::FindInContainer(IUnknown*  pContainer,  LONG nSearchFlags,
                                              BSTR       bstrTagName, SAFEARRAY* psa,
                                              IUnknown** ppResult)
{
	if ((NULL == bstrTagName) && !(nSearchFlags & Common::SEARCH_FRAME) &&
		!(nSearchFlags & Common::SEARCH_HTML_DIALOG) && !(nSearchFlags & Common::SEARCH_MODAL_HTML_DLG))
	{
		traceLog << "bstrTagName is null in CExplorerPlugin::FindInContainer\n";
		return E_INVALIDARG;
	}

	if (NULL == psa)
	{
		traceLog << "psa is null in CExplorerPlugin::FindInContainer\n";
		return E_INVALIDARG;
	}

	if (NULL == ppResult)
	{
		traceLog << "ppResult is null in CExplorerPlugin::FindInContainer\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IWebBrowser2> spBrws;
	CComQIPtr<IHTMLWindow2> spFrame;
	CComQIPtr<IHTMLElement> spElement;

	if (NULL == pContainer)
	{
		// pContainer is NULL when the container is the browser. Get the IWebBrowser2 pointer.
		HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
		if (spBrws == NULL)
		{
			traceLog << "Can not get the IWebBrowser2 in CExplorerPlugin::FindInContainer\n";
			return E_FAIL;
		}
	}
	else
	{
		spFrame = pContainer;
		if (spFrame == NULL)
		{
			spElement = pContainer;
			if (spElement == NULL)
			{
				traceLog << "Invalid pContainer input parameter in CExplorerPlugin::FindInContainer\n";
				return E_INVALIDARG;
			}
		}
	}


	// Search the element in browser.
	try
	{
		CComQIPtr<IUnknown, &IID_IUnknown> spResult;

		if (spBrws != NULL)
		{
			if (nSearchFlags & Common::SEARCH_COLLECTION)
			{
				ATLASSERT(!(nSearchFlags & SEARCH_FRAME)); // Frame collection is computed differently.
				spResult = FindElementAlgorithms::FindAllElements(spBrws, bstrTagName, psa);
			}
			else
			{
				if (nSearchFlags & SEARCH_FRAME)
				{
					spResult = FindFrameAlgorithms::FindFrame(spBrws, psa);
				}
				else if (nSearchFlags & SEARCH_HTML_DIALOG)
				{
					spResult = FindFrameAlgorithms::FindHtmlDialog(spBrws, psa);
				}
				else if (nSearchFlags & SEARCH_MODAL_HTML_DLG)
				{
					spResult = FindFrameAlgorithms::FindModalHtmlDialog(spBrws, psa);
				}
				else
				{
					spResult = FindElementAlgorithms::FindElement(spBrws, bstrTagName, psa);
				}
			}
		}
		else if (spFrame != NULL)
		{
			ATLASSERT(!(nSearchFlags & SEARCH_HTML_DIALOG));
			ATLASSERT(!(nSearchFlags & SEARCH_MODAL_HTML_DLG));

			if (nSearchFlags & Common::SEARCH_COLLECTION)
			{
				ATLASSERT(!(nSearchFlags & SEARCH_FRAME)); // Frame collection is computed differently.
				if (nSearchFlags & Common::SEARCH_CHILDREN_ONLY)
				{
					spResult = FindElementAlgorithms::FindChildrenElements(spFrame, bstrTagName, psa);
				}
				else
				{
					spResult = FindElementAlgorithms::FindAllElements(spFrame, bstrTagName, psa);
				}
			}
			else
			{
				if (nSearchFlags & Common::SEARCH_CHILDREN_ONLY)
				{
					if (nSearchFlags & SEARCH_FRAME)
					{
						spResult = FindFrameAlgorithms::FindChildFrame(spFrame, psa);
					}
					else
					{
						spResult = FindElementAlgorithms::FindChildElement(spFrame, bstrTagName, psa);
					}
				}
				else
				{
					if (nSearchFlags & SEARCH_FRAME)
					{
						spResult = FindFrameAlgorithms::FindFrame(spFrame, psa);
					}
					else
					{
						spResult = FindElementAlgorithms::FindElement(spFrame, bstrTagName, psa);
					}
				}
			}
		}
		else
		{
			ATLASSERT(!(nSearchFlags & SEARCH_FRAME)); // Don't serach frames inside elements.
			ATLASSERT(spElement != NULL);

			if (nSearchFlags & Common::SEARCH_COLLECTION)
			{
				if (nSearchFlags & Common::SEARCH_CHILDREN_ONLY)
				{
					spResult = FindElementAlgorithms::FindChildrenElements(spElement, bstrTagName, psa);
				}
				else
				{
					spResult = FindElementAlgorithms::FindAllElements(spElement, bstrTagName, psa);
				}
			}
			else
			{
				if (nSearchFlags & Common::SEARCH_CHILDREN_ONLY)
				{
					spResult = FindElementAlgorithms::FindChildElement(spElement, bstrTagName, psa);
				}
				else if (nSearchFlags & Common::SEARCH_PARENT_ELEM)
				{
					spResult = FindElementAlgorithms::FindParentElement(spElement, bstrTagName, psa);
				}
				else
				{
					spResult = FindElementAlgorithms::FindElement(spElement, bstrTagName, psa);
				}
			}
		}

		// If the result collection is empty re-try in searchTimeout.
		CComQIPtr<ILocalElementCollection> spColl = spResult;
		if (spColl != NULL)
		{
			BOOL bIsEmpty;
			if (SUCCEEDED(IsEmptyCollection(spColl, bIsEmpty)))
			{
				if (bIsEmpty)
				{
					*ppResult = spResult.Detach();
					return S_FALSE;
				}
			}
			else
			{
				return E_FAIL;
			}
		}

		*ppResult = spResult.Detach();
		if (NULL == *ppResult)
		{
			return S_FALSE;
		}
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		return E_INVALIDARG;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::FindFramesInContainer(IUnknown*       pContainer,
                                                    LONG            nSearchFlags,
                                                    SAFEARRAY*      psa,
                                                    LONG*           pSize,
                                                    IHTMLWindow2*** pppFrames)
{
	if (NULL == psa)
	{
		traceLog << "psa is null in CExplorerPlugin::FindFramesInContainer\n";
		return E_INVALIDARG;
	}

	if (NULL == pSize)
	{
		traceLog << "pSize is null in CExplorerPlugin::FindFramesInContainer\n";
		return E_INVALIDARG;
	}

	if (NULL == pppFrames)
	{
		traceLog << "pppFrames is null in CExplorerPlugin::FindFramesInContainer\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IWebBrowser2> spBrws;
	CComQIPtr<IHTMLWindow2> spFrame;

	if (NULL == pContainer)
	{
		// pContainer is NULL when the container is the browser. Get the IWebBrowser2 pointer.
		HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
		if (spBrws == NULL)
		{
			traceLog << "Can not get the IWebBrowser2 in CExplorerPlugin::FindFramesInContainer\n";
			return E_FAIL;
		}
	}
	else
	{
		spFrame = pContainer;
		if (spFrame == NULL)
		{
			traceLog << "Invalid pContainer input parameter in CExplorerPlugin::FindFramesInContainer\n";
			return E_INVALIDARG;
		}
	}

	try
	{
		list<CAdapt<CComQIPtr<IHTMLWindow2> > > frameList;
		if (spBrws != NULL)
		{
			ATLASSERT(spFrame == NULL);
			ATLASSERT(!(nSearchFlags & Common::SEARCH_CHILDREN_ONLY));	// There is no Browser.FindChildrenFrames method !
			frameList = FindFrameAlgorithms::FindAllFrames(spBrws, psa, nSearchFlags);
		}
		else
		{
			ATLASSERT(spFrame != NULL);
			if (nSearchFlags & Common::SEARCH_CHILDREN_ONLY)
			{
				frameList = FindFrameAlgorithms::FindChildrenFrames(spFrame, psa, nSearchFlags);
			}
			else
			{
				frameList = FindFrameAlgorithms::FindAllFrames(spFrame, psa, nSearchFlags);
			}
		}

		if (frameList.size() == 0)
		{
			// Collection is empty. Do nothing.
			return S_FALSE;
		}

		LPVOID         lpMemCollection = ::CoTaskMemAlloc(sizeof(IHTMLWindow2*) * frameList.size());
		IHTMLWindow2** ppFrames        = static_cast<IHTMLWindow2**>(lpMemCollection);
		if (NULL == ppFrames)
		{
			traceLog << "CoTaskMemAlloc returned NULL in CExplorerPlugin::FindFramesInContainer\n";
			return E_FAIL;
		}

		*pSize = static_cast<LONG>(frameList.size());

		int i = 0;
		for (list<CAdapt<CComQIPtr<IHTMLWindow2> > >::iterator it = frameList.begin();
			 it != frameList.end(); ++it)
		{
			ppFrames[i++] = (*it).m_T.Detach();
		}

		ATLASSERT(NULL == *pppFrames);
		*pppFrames = ppFrames;
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		return E_INVALIDARG;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetScreenClickPoint(IHTMLElement* pElement, LONG nGetInputFileButton, LONG* pX, LONG* pY)
{
	if ((NULL == pElement) || (NULL == pX) || (NULL == pY))
	{
		traceLog << "Invalid argument in CExplorerPlugin::ClickHtmlElement\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLElement> spElementToScrollIntoView;
	CComQIPtr<IHTMLAreaElement> spArea = pElement;
	if (spArea != NULL)
	{
		spElementToScrollIntoView = HtmlHelpers::GetAreaImage(spArea);
		if (spElementToScrollIntoView == NULL)
		{
			traceLog << "HtmlHelpers::GetAreaImage failed in CExplorerPlugin::ClickHtmlElement\n";
			return E_FAIL;
		}
	}
	else
	{
		spElementToScrollIntoView = pElement;
	}

	// Scroll the element into view.
	HRESULT hRes = spElementToScrollIntoView->scrollIntoView(CComVariant(TRUE));
	if (FAILED(hRes))
	{
		traceLog << "IHTMLElement::scrollIntoView failed with code " << hRes << " in CExplorerPlugin::ClickHtmlElement\n";
		return hRes;
	}

	// Get the click point.
	POINT pClickScreenPoint = { 0 };
	int   nZoomLevel = GetZoomLevel();
	BOOL  bRes       = HtmlHelpers::GetElemClickScreenPoint(spElementToScrollIntoView, &pClickScreenPoint, nGetInputFileButton, nZoomLevel);

	if (!bRes)
	{
		traceLog << "GetElemClickScreenPoint failed in CExplorerPlugin::ClickHtmlElement\n";
		return E_FAIL;
	}

	*pX = pClickScreenPoint.x;
	*pY = pClickScreenPoint.y;

	if (spArea != NULL)
	{
		POINT offsetPoint;
		BOOL bRes = HtmlHelpers::GetAreaOffset(spArea, &offsetPoint);
		if (!bRes)
		{
			traceLog << "HtmlHelpers::GetAreaOffset failed in CExplorerPlugin::GetScreenClickPoint\n";
			return E_FAIL;
		}

		*pX += ((offsetPoint.x * nZoomLevel) / 100);
		*pY += ((offsetPoint.y * nZoomLevel) / 100);
	}

	return S_OK;
}

HRESULT CExplorerPlugin::FindInfoForSelect
	(
		IHTMLElement* pElement,
		VARIANT       vStart,
		VARIANT       vEnd,
		LONG          nFlags,
		
		list<DWORD>& outItemsToSelect,
		BOOL&        bOutIsCombo,
		BOOL&        bOutMultiple
	)
{
	// Check parameter.
	if (!Common::IsValidOptionVariant(vStart) ||
	    ((VT_EMPTY != vEnd.vt) && !Common::IsValidOptionVariant(vEnd)))
	{
		traceLog << "Invalid parameters in CExplorerPlugin::FindInfoForSelect\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLSelectElement> spSelect = pElement;
	if (spSelect == NULL)
	{
		traceLog << "pElement is not a <select> html element in CExplorerPlugin::FindInfoForSelect\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}

	BOOL bIsCombo  = HtmlHelpers::IsComboControl(spSelect);
	BOOL bMultiple = FALSE;
	if (!bIsCombo)
	{
		VARIANT_BOOL vbMultiple;
		HRESULT      hRes = spSelect->get_multiple(&vbMultiple);
		if (FAILED(hRes))
		{
			traceLog << "IHTMLSelectElement::get_multiple failed with code " << hRes << " in CExplorerPlugin::FindInfoForSelect\n";
			return E_FAIL;
		}

		if (VARIANT_TRUE == vbMultiple)
		{
			bMultiple = TRUE;
		}
		else
		{
			ATLASSERT(VARIANT_FALSE == vbMultiple);
			bMultiple = FALSE;
		}
	}

	if (!bMultiple && ((nFlags & Common::ADD_SELECTION) || (VT_EMPTY != vEnd.vt)))
	{
		traceLog << "Can not add selection for combo or simple selection list in CExplorerPlugin::FindInfoForSelect\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}


	// Find the items to select.
	try
	{
		outItemsToSelect.clear();
		outItemsToSelect = FindItemsToSelect(vStart, vEnd, pElement, nFlags);
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		return E_INVALIDARG;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	if (outItemsToSelect.size() == 0)
	{
		traceLog << "Options not found in CExplorerPlugin::FindInfoForSelect\n";
		return HRES_NOT_FOUND_ERR;
	}

	bOutIsCombo  = bIsCombo;
	bOutMultiple = bMultiple;

	return S_OK;
}


HRESULT CExplorerPlugin::DoSelectOptions(IHTMLElement* pElement, const list<DWORD>& itemsToSelect, BOOL bIsCombo, BOOL bMultiple, LONG nFlags)
{
	if (!pElement)
	{
		traceLog << "pElement is NULL in CExplorerPlugin::DoSelectOptions\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLSelectElement> spSelect = pElement;
	if (!spSelect)
	{
		traceLog << "Not a IHTMLSelectElement in CExplorerPlugin::DoSelectOptions\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}

	try
	{
		if (bIsCombo)
		{
			if (!SelectComboItem(spSelect, *itemsToSelect.begin()))
			{
				traceLog << "Index out of bound for SelectComboItem in CExplorerPlugin::DoSelectOptions\n";
				return HRES_INDEX_OUT_OF_BOUND_ERR;
			}
		}
		else if (!bMultiple)
		{
			if (!SelectSingleListItem(spSelect, *itemsToSelect.begin()))
			{
				traceLog << "Index out of bound for SelectListItem in CExplorerPlugin::DoSelectOptions\n";
				return HRES_INDEX_OUT_OF_BOUND_ERR;
			}
		}
		else
		{
			if (!(nFlags & Common::ADD_SELECTION))
			{
				if  (!SelectMultipleListItem(spSelect, itemsToSelect))
				{
					traceLog << "Index out of bound for AddToSelection in CExplorerPlugin::DoSelectOptions\n";
					return HRES_INDEX_OUT_OF_BOUND_ERR;
				}
			}
			else if (!AddToSelection(spSelect, itemsToSelect))
			{
				traceLog << "Index out of bound for AddToSelection in CExplorerPlugin::DoSelectOptions\n";
				return HRES_INDEX_OUT_OF_BOUND_ERR;
			}
		}
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::SelectOptions(IHTMLElement* pElement, VARIANT vStart, VARIANT vEnd, LONG nFlags)
{

	list<DWORD> itemsToSelect;
	BOOL bIsCombo  = FALSE;
	BOOL bMultiple = FALSE;

	HRESULT hRes = CExplorerPlugin::FindInfoForSelect(pElement, vStart, vEnd, nFlags, itemsToSelect, bIsCombo, bMultiple);
	if (FAILED(hRes))
	{
		traceLog << "FindInfoForSelect failed in CExplorerPlugin::SelectOptions\n";
		return hRes;
	}

	if (Common::PERFORM_ASYNC_ACTION & nFlags)
	{
		SelectAsyncAction* pSelectAsync = new SelectAsyncAction(this, pElement, itemsToSelect, bIsCombo, bMultiple, nFlags);

		// Put the action object into the queue and post a message.
		m_asyncActionsQueue.push_back(pSelectAsync);
		this->PostMessage(WM_APP_ASYNC_ACTION);

		return S_OK;
	}
	else
	{
		return DoSelectOptions(pElement, itemsToSelect, bIsCombo, bMultiple, nFlags);
	}
}

BOOL CExplorerPlugin::SelectMultipleListItem(CComQIPtr<IHTMLSelectElement> spSelectElem, const list<DWORD>& itemsToSelect)
{
	ATLASSERT(spSelectElem != NULL);
	ATLASSERT(itemsToSelect.size() > 0);

	// TODO: assume that inside <select> there are ONLY <options>.
	LONG    nOptionsCount = 0;
	HRESULT hRes          = spSelectElem->get_length(&nOptionsCount);
	if (FAILED(hRes))
	{
		traceLog << "get_length failed in CExplorerPlugin::SelectMultipleListItem with code: " << hRes << "\n";
		throw CreateException(_T("get_length failed in CExplorerPlugin::SelectMultipleListItem"));
	}

	BOOL bCurrentSelectionExactMatch = TRUE;
	std::set<DWORD> setOfItemToSelect(itemsToSelect.begin(), itemsToSelect.end());

	for (LONG i = 0; i < nOptionsCount; ++i)
	{
		CComVariant          varIndex(i);
		CComQIPtr<IDispatch> spDispOption;
		hRes = spSelectElem->item(varIndex, CComVariant(0), &spDispOption);

		CComQIPtr<IHTMLOptionElement> spCrntOption = spDispOption;
		if (FAILED(hRes) || (spCrntOption == NULL))
		{
			throw CreateException(_T("IHTMLSelectElement::item failed in CExplorerPlugin::SelectMultipleListItem"));
		}

		// In case of an error consider the item not selected.
		VARIANT_BOOL vbIsSelected = VARIANT_FALSE;
		hRes = spCrntOption->get_selected(&vbIsSelected);

		if (vbIsSelected != VARIANT_FALSE)
		{
			if (0 == setOfItemToSelect.erase(i))
			{
				bCurrentSelectionExactMatch = FALSE;
				break;
			}
		}
	}

	bCurrentSelectionExactMatch = (bCurrentSelectionExactMatch && setOfItemToSelect.empty());
	if (!bCurrentSelectionExactMatch)
	{
		ClearSelection(spSelectElem, FALSE);
		return AddToSelection(spSelectElem, itemsToSelect);
	}
	else
	{
		// The current selection exactly match itemsToSelect. Don't do anything.
		return TRUE;
	}
}


BOOL CExplorerPlugin::AddToSelection(CComQIPtr<IHTMLSelectElement> spSelectElem, const list<DWORD>& itemsToSelect)
{
	ATLASSERT(spSelectElem != NULL);
	ATLASSERT(itemsToSelect.size() > 0);

	// TODO: assume that inside <select> there are ONLY <options>.
	LONG    nOptionsCount = 0;
	HRESULT hRes          = spSelectElem->get_length(&nOptionsCount);
	if (FAILED(hRes))
	{
		traceLog << "get_length failed in CExplorerPlugin::AddToSelection with code: " << hRes << "\n";
		throw CreateException(_T("get_length failed in CExplorerPlugin::AddToSelection"));
	}

	BOOL bGenerateOnChange = FALSE;
	for (list<DWORD>::const_iterator it = itemsToSelect.begin();
	     it != itemsToSelect.end(); ++it)
	{
		DWORD dwItem = *it;
		if ((DWORD)nOptionsCount <= dwItem)
		{
			traceLog << "Index:" << dwItem << "out of bound: " << nOptionsCount << " in CExplorerPlugin::AddToSelection\n";
			return FALSE;
		}

		CComVariant          varIndex(dwItem);
		CComQIPtr<IDispatch> spDispOption;
		hRes = spSelectElem->item(varIndex, CComVariant(0), &spDispOption);

		CComQIPtr<IHTMLOptionElement> spCrntOption = spDispOption;
		if (FAILED(hRes) || (spCrntOption == NULL))
		{
			throw CreateException(_T("IHTMLSelectElement::item failed in CExplorerPlugin::AddToSelection"));
		}

		// In case of an error consider the item not selected.
		VARIANT_BOOL vbIsAlreadySelected = VARIANT_FALSE;
		hRes = spCrntOption->get_selected(&vbIsAlreadySelected);

		if (VARIANT_FALSE == vbIsAlreadySelected)
		{
			bGenerateOnChange = TRUE; // At least one item was added to selection.

			// The item is not already selected.
			hRes = spCrntOption->put_selected(VARIANT_TRUE);
			if (FAILED(hRes))
			{
				throw CreateException(_T("IHTMLOptionElement::put_selected failed in CExplorerPlugin::AddToSelection"));
			}
		}
	}

	if (bGenerateOnChange)
	{
		BOOL bRes = HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onchange");
		if (!bRes)
		{
			throw CreateException(_T("HtmlHelpers::FireEventOnElement failed in CExplorerPlugin::AddToSelection"));
		}
	}

	return TRUE;
}


// Returns FALSE if dwItem is greater or equal than the number of <options> inside <select>.
// Throws exception for any other error.
BOOL CExplorerPlugin::SelectComboItem(CComQIPtr<IHTMLSelectElement> spSelectElem, DWORD dwItem)
{
	ATLASSERT(spSelectElem != NULL);
	ATLASSERT(HtmlHelpers::IsComboControl(spSelectElem));

	// TODO: assume that inside <select> there are ONLY <options>.
	LONG    nOptionsCount = 0;
	HRESULT hRes          = spSelectElem->get_length(&nOptionsCount);
	if (FAILED(hRes))
	{
		traceLog << "get_length failed in CExplorerPlugin::SelectComboItem with code: " << hRes << "\n";
		throw CreateException(_T("get_length failed in CExplorerPlugin::SelectComboItem"));
	}

	if ((DWORD)nOptionsCount <= dwItem)
	{
		traceLog << "Index:" << dwItem << "out of bound: " << nOptionsCount << " in CExplorerPlugin::SelectComboItem\n";
		return FALSE;
	}

	// On failure assume there is no selection.
	LONG nCrntSelectedIndex = -1;
	hRes = spSelectElem->get_selectedIndex(&nCrntSelectedIndex);

	// Don't do anything if already selected .
	if (nCrntSelectedIndex != dwItem)
	{
		BOOL bRes = HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onbeforeactivate");

		// Select the item with index dwItem.
		hRes = spSelectElem->put_selectedIndex(dwItem);
		if (FAILED(hRes))
		{
			traceLog << "put_selectedIndex failed in CExplorerPlugin::SelectComboItem with code: " << hRes << "\n";
			throw CreateException(_T("put_selectedIndex failed in CExplorerPlugin::SelectComboItem"));
		}

		// The following sequence of events are needed for jQuery onchange processing.
		bRes = bRes && HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onchange");
		bRes = bRes && HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onclick");

		if (!bRes)
		{
			throw CreateException(_T("HtmlHelpers::FireEvent(s)OnElement failed in CExplorerPlugin::SelectComboItem"));
		}
	}

	return TRUE;
}


BOOL CExplorerPlugin::SelectSingleListItem(CComQIPtr<IHTMLSelectElement> spSelectElem, DWORD dwItem)
{
	ATLASSERT(spSelectElem != NULL);
	ATLASSERT(!HtmlHelpers::IsComboControl(spSelectElem));
	ATLASSERT(!HtmlHelpers::IsMultipleListControl(spSelectElem));

	// TODO: assume that inside <select> there are ONLY <options>.
	LONG    nOptionsCount = 0;
	HRESULT hRes          = spSelectElem->get_length(&nOptionsCount);
	if (FAILED(hRes))
	{
		traceLog << "get_length failed in CExplorerPlugin::SelectSingleListItem with code: " << hRes << "\n";
		throw CreateException(_T("get_length failed in CExplorerPlugin::SelectSingleListItem"));
	}

	if ((DWORD)nOptionsCount <= dwItem)
	{
		traceLog << "Index:" << dwItem << "out of bound: " << nOptionsCount << " in CExplorerPlugin::SelectSingleListItem\n";
		return FALSE;
	}

	// On failure assume there is no selection.
	LONG nCrntSelectedIndex = -1;
	hRes = spSelectElem->get_selectedIndex(&nCrntSelectedIndex);

	// Don't do anything if already selected .
	if (nCrntSelectedIndex != dwItem)
	{
		// Select the item with index dwItem.
		hRes = spSelectElem->put_selectedIndex(dwItem);
		if (FAILED(hRes))
		{
			traceLog << "put_selectedIndex failed in CExplorerPlugin::SelectSingleListItem with code: " << hRes << "\n";
			throw CreateException(_T("put_selectedIndex failed in CExplorerPlugin::SelectSingleListItem"));
		}

		BOOL bRes = HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onchange");
		if (!bRes)
		{
			throw CreateException(_T("HtmlHelpers::FireEventOnElement failed in CExplorerPlugin::SelectSingleListItem"));
		}
	}

	return TRUE;
}


void CExplorerPlugin::ClearSelection(CComQIPtr<IHTMLSelectElement> spSelectElem, BOOL bFireEvent)
{
	ATLASSERT(spSelectElem != NULL);

	// On error assume there's no selection.
	LONG nSelectedIndex = -1;
	HRESULT hRes = spSelectElem->get_selectedIndex(&nSelectedIndex);

	if (nSelectedIndex != -1)
	{
		hRes = spSelectElem->put_selectedIndex(-1);
		if (FAILED(hRes))
		{
			throw CreateException(_T("put_selectedIndex failed in CExplorerPlugin::ClearSelection"));
		}

		if (bFireEvent)
		{
			BOOL bRes = HtmlHelpers::FireEventOnElement(spSelectElem.p, L"onchange");
			if (!bRes)
			{
				throw CreateException(_T("HtmlHelpers::FireEventOnElement failed in CExplorerPlugin::ClearSelection"));
			}
		}
	}
}


list<DWORD> CExplorerPlugin::FindItemsToSelect(VARIANT vItems, IHTMLElement* pElement, LONG nFlags)
{
	ATLASSERT(Common::IsValidOptionVariant(vItems));

	// Find the items to select.
	list<DWORD> itemsToSelect;
	if (vItems.vt != VT_BSTR)
	{
		DWORD dwItem;
		switch (vItems.vt)
		{
			case VT_I4:  dwItem = vItems.lVal;  break;
			case VT_I2:  dwItem = vItems.iVal;  break;
			case VT_UI4: dwItem = vItems.ulVal; break;
			case VT_UI2: dwItem = vItems.uiVal; break;
			default: ATLASSERT(FALSE);
		};

		itemsToSelect.push_back(dwItem);
	}
	else
	{
		itemsToSelect = FindElementAlgorithms::FindOptionsIndexes(pElement, vItems.bstrVal);
	}

	return itemsToSelect;
}


list<DWORD> CExplorerPlugin::FindItemsToSelect(VARIANT vStart, VARIANT vEnd, IHTMLElement* pElement, LONG nFlags)
{
	ATLASSERT((Common::IsValidOptionVariant(vStart) ||
	          ((VT_EMPTY != vEnd.vt) && Common::IsValidOptionVariant(vEnd))));

	if (VT_EMPTY == vEnd.vt)
	{
		return FindItemsToSelect(vStart, pElement, nFlags);
	}

	list<DWORD> startList = FindItemsToSelect(vStart, pElement, nFlags);
	list<DWORD> endList   = FindItemsToSelect(vEnd,   pElement, nFlags);

	if ((startList.size() == 0) || (endList.size() == 0))
	{
		return list<DWORD>();
	}

	DWORD dwStart = *(startList.begin());
	DWORD dwEnd   = *(endList.rbegin());
	if (dwStart > dwEnd)
	{
		return list<DWORD>();
	}

	CComQIPtr<IHTMLSelectElement> spSelectElem = pElement;
	if (spSelectElem == NULL)
	{
		throw CreateException(_T("IHTMLSelectElement not available in CExplorerPlugin::FindItemsToSelect"));
	}

	LONG    nOptionsCount = 0;
	HRESULT hRes          = spSelectElem->get_length(&nOptionsCount);
	if (FAILED(hRes))
	{
		traceLog << "get_length failed in CExplorerPlugin::FindItemsToSelect with code: " << hRes << "\n";
		throw CreateException(_T("get_length failed in CExplorerPlugin::FindItemsToSelect"));
	}

	list<DWORD> resultList;

	if (dwStart >= (DWORD)nOptionsCount)
	{
		// AddToSelection will eventually report an index out of bound condition.
		resultList.push_back(dwStart);
		return resultList;
	}

	if (dwEnd >= (DWORD)nOptionsCount)
	{
		// AddToSelection will eventually report an index out of bound condition.
		resultList.push_back(dwEnd);
		return resultList;
	}

	for (DWORD dwIndex = dwStart; dwIndex <= dwEnd; ++dwIndex)
	{
		resultList.push_back(dwIndex);
	}

	return resultList;
}


STDMETHODIMP CExplorerPlugin::ClearSelection(IHTMLElement* pElement)
{
	if (NULL == pElement)
	{
		traceLog << "Invalid parameters in CExplorerPlugin::ClearSelection\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLSelectElement> spSelect = pElement;
	if (spSelect == NULL)
	{
		traceLog << "pElement is not a <select> html element in CExplorerPlugin::ClearSelection\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}

	try
	{
		ClearSelection(spSelect, TRUE);
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CExplorerPlugin::GetElementText(IHTMLElement* pElement, BSTR* pBstrText)
{
	if ((NULL == pElement) || (NULL == pBstrText))
	{
		traceLog << "Invald parameters in CExplorerPlugin::GetElementText\n";
		return E_INVALIDARG;
	}

	CComBSTR bstrOutText = HtmlHelpers::GetTextAttributeValue(pElement);
	*pBstrText = bstrOutText.Detach();

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetIEServerWnd(IHTMLElement* pElement, LONG* pWnd)
{
	if (NULL == pWnd)
	{
		traceLog << "Invald parameters in CExplorerPlugin::GetIEServerWnd\n";
		return E_INVALIDARG;
	}

	HRESULT hRes;
	CComQIPtr<IHTMLDocument2> spDocument;

	if (NULL == pElement)
	{
		// Get the IWebBrowser2 pointer.
		CComQIPtr<IWebBrowser2> spBrws;
		GetSite(IID_IWebBrowser2, (void**)&spBrws);
		if (spBrws == NULL)
		{
			traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::GetIEServerWnd\n";
			return E_FAIL;
		}

		if (Common::GetIEVersion() > 6)
		{
			HWND   hTabWnd   = HtmlHelpers::GetIEWndFromBrowser(spBrws);
			String sTabClass = Common::GetWndClass(hTabWnd);

			if (!::IsWindow(hTabWnd) || ((sTabClass != _T("TabWindowClass") && (sTabClass != _T("Shell Embedding")))))
			{
				traceLog << "Can NOT obtain TabWindowClass wnd in CExplorerPlugin::GetIEServerWnd\n";
				return E_FAIL;
			}

			HWND hIEServWnd = Common::GetChildWindowByClassName(hTabWnd, _T("Internet Explorer_Server"));
			if (hIEServWnd != NULL)
			{
				*pWnd = HandleToLong(hIEServWnd);
				return S_OK;
			}
			else
			{
				traceLog << "Can NOT obtain 'Internet Explorer_Server' wnd in CExplorerPlugin::GetIEServerWnd\n";
				return E_FAIL;
			}
		}
		else
		{
			CComQIPtr<IDispatch> spDisp;
			hRes = spBrws->get_Document(&spDisp);
			spDocument = spDisp;
		}
	}
	else
	{
		// Get the document of the element.
		CComQIPtr<IDispatch> spDisp;
		hRes = pElement->get_document(&spDisp);
		spDocument = spDisp;
	}

	if (FAILED(hRes) || (spDocument == NULL))
	{
		traceLog << "Can not get the document in CExplorerPlugin::GetIEServerWnd\n";
		return E_FAIL;
	}

	// Get the body element.
	CComQIPtr<IHTMLElement> spBody;
	hRes = spDocument->get_body(&spBody);
	if (FAILED(hRes) || (spBody == NULL))
	{
		traceLog << "Can not get the body element in CExplorerPlugin::GetIEServerWnd\n";
		return E_FAIL;
	}

	HWND hIeServerWnd = NULL;
	hRes = HtmlHelpers::GetIEServerWndFromElement(spBody, &hIeServerWnd);

	*pWnd = HandleToLong(hIeServerWnd);
	return hRes;
}


STDMETHODIMP CExplorerPlugin::Invoke
				(DISPID dispidMember, REFIID riid, LCID lcid, 
				 WORD   nFlags, DISPPARAMS* pDispParams,
				 VARIANT* pvarResult, EXCEPINFO*  pExcepInfo, UINT* puArgErr)
{
	if (NULL == pDispParams)
	{
		traceLog << "Invalid parameter in CExplorerPlugin::Invoke\n";
		return E_INVALIDARG;
	}

	switch (dispidMember)
	{
		case DISPID_NAVIGATEERROR:
		{
			traceLog << "DISPID_NAVIGATEERROR in CExplorerPlugin::Invoke\n";
			CComVariant	vErr = *(pDispParams->rgvarg[1].pvarVal);
			ATLASSERT(VT_I4 == pDispParams->rgvarg[1].pvarVal->vt);

			m_nLastNavigationErr = vErr.lVal;
			break;
		}

		case DISPID_BEFORENAVIGATE2:
		{
			traceLog << "DISPID_BEFORENAVIGATE2 in CExplorerPlugin::Invoke\n";

			// Reset the last navigation error.
			m_nLastNavigationErr = 0;
			break;
		}

		case DISPID_DOWNLOADBEGIN:
		{
			traceLog << "DISPID_DOWNLOADBEGIN in CExplorerPlugin::Invoke m_nActiveDownloads=" << m_nActiveDownloads << "\n";

			ATLASSERT(m_nActiveDownloads >= 0);
			m_nActiveDownloads++;
			break;
		}

		case DISPID_DOWNLOADCOMPLETE:
		{
			traceLog << "DISPID_DOWNLOADCOMPLETE in CExplorerPlugin::Invoke m_nActiveDownloads=" << m_nActiveDownloads << "\n";
			if (m_nActiveDownloads > 0)
			{
				m_nActiveDownloads--;
			}

			break;
		}

		case DISPID_ONQUIT:
		{
			OnQuit();
			return S_OK;
		}
	}

	return IDispatchImpl<IExplorerPlugin, &IID_IExplorerPlugin, &LIBID_OpenTwebstPluginLib>::Invoke(dispidMember, riid, lcid, 
	                                                                                  nFlags, pDispParams, pvarResult,
                                                                                      pExcepInfo, puArgErr);
}


void CExplorerPlugin::OnQuit()
{
	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);

	if (spBrws != NULL)
	{
		// Unregister to stop receiving browser events.
		if (m_dwEventsCookie != 0)
		{
			hRes = AtlUnadvise(spBrws, DIID_DWebBrowserEvents2, m_dwEventsCookie);
			if (FAILED(hRes))
			{
				traceLog << "Failed to unregister for browser events in CExplorerPlugin::OnQuit, code=" << hRes << "\n";
			}
		}
	}
	else
	{
		traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::OnQuit\n";
	}

	m_dwEventsCookie = 0;

	if (m_bIsForceLoaded)
	{
		UnInit();

		// Auto-delete object if it was hook-injected rather than loaded by IE as a BHO.
		Release();
	}
}


STDMETHODIMP CExplorerPlugin::GetLastNavigationErr(LONG* pErrCode)
{
	if (NULL == pErrCode)
	{
		traceLog << "Invalid parameter in CExplorerPlugin::GetLastNavigationErr\n";
		return E_INVALIDARG;
	}

	*pErrCode = m_nLastNavigationErr;
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetElementScreenRect(IHTMLElement* pElement, long* pLeft, long* pTop, long* pRight, long* pBottom)
{
	// Check parameters.
	if ((NULL == pElement) || (NULL == pLeft) || (NULL == pTop) || (NULL == pRight) || (NULL == pBottom))
	{
		traceLog << "Invalid parameters in CExplorerPlugin::GetElementScreenRect\n";
		return E_INVALIDARG;
	}

	// Get the IWebBrowser2 pointer.
	int                     nZoomLevel = 100;
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT                 hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);

	if (SUCCEEDED(hRes) && spBrws)
	{
		// It seems the call below does NOT work from outside IE.
		CComVariant vZoom;
		hRes = spBrws->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, NULL, &vZoom);

		if (SUCCEEDED(hRes))
		{
			nZoomLevel = vZoom.lVal;
		}
	}

	RECT screenRect = { 0 };
	BOOL bRes = HtmlHelpers::GetElementScreenLocation(pElement, &screenRect, nZoomLevel);
	if (!bRes)
	{
		traceLog << "HtmlHelpers::GetElementScreenLocation failed in CExplorerPlugin::GetElementScreenRect\n";
		return E_FAIL;
	}

	*pLeft	 = screenRect.left;
	*pTop	 = screenRect.top;
	*pRight  = screenRect.right;
	*pBottom = screenRect.bottom;

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetDocumentFromWindow(IHTMLWindow2* pWindow, IHTMLDocument2** ppDocument)
{
	if ((NULL == pWindow) || (NULL == ppDocument))
	{
		traceLog << "Invalid parameters in CExplorerPlugin::GetDocumentFromWindow\n";
		return E_INVALIDARG;
	}

	// Get the document from the window.
	CComQIPtr<IHTMLDocument2> spDoc = HtmlHelpers::HtmlWindowToHtmlDocument(pWindow);

	*ppDocument = spDoc.Detach();
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetAppName(BSTR* pBstrApp)
{
	if (NULL == pBstrApp)
	{
		traceLog << "Invalid parameters in CExplorerPlugin::GetAppName\n";
		return E_INVALIDARG;
	}

	HMODULE hExe = ::GetModuleHandle(NULL);
	if (NULL == hExe)
	{
		traceLog << "GetModuleHandle failed in CExplorerPlugin::GetAppName\n";
		return E_FAIL;
	}

	TCHAR szFullPath[MAX_PATH + 1] = { 0 };
	DWORD dwRes = ::GetModuleFileName(hExe, szFullPath, sizeof(szFullPath) / sizeof(szFullPath[0]) - 1);
	if ((0 == dwRes) || (ERROR_INSUFFICIENT_BUFFER == ::GetLastError()))
	{
		traceLog << "GetModuleFileName failed in CExplorerPlugin::GetAppName\n";
		return E_FAIL;
	}

	CPath path(szFullPath);
	path.StripPath();

	CComBSTR bstrAppName = path.m_strPath.MakeLower();
	*pBstrApp = bstrAppName.Detach();

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetBrowserTitle(BSTR* pBstrTitle)
{
	if (NULL == pBstrTitle)
	{
		traceLog << "Invalid parameters in CExplorerPlugin::GetBrowserTitle\n";
		return E_INVALIDARG;
	}

	try
	{
		String sTitle;
		BOOL   bRes = GetBrowserTitle(&sTitle);
		if (!bRes)
		{
			traceLog << "GetBrowserTitle(String*) failed\n";
			return E_FAIL;
		}

		*pBstrTitle = CComBSTR(sTitle.c_str()).Detach();
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		traceLog << "GetBrowserTitle failed\n";
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetBrowserThreadID(LONG* pThID)
{
	if (NULL == pThID)
	{
		return E_INVALIDARG;
	}

	*pThID = ::GetCurrentThreadId();
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetBrowserProcessID(LONG* pProcID)
{
	if (NULL == pProcID)
	{
		return E_INVALIDARG;
	}

	*pProcID = ::GetCurrentProcessId();
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::SetFocusOnElement(IHTMLElement* pTargetElement, VARIANT_BOOL vbAsync)
{
	if (VARIANT_FALSE == vbAsync)
	{
		SetFocusAsyncAction setFocusAction(pTargetElement);
		return setFocusAction.DoAction();
	}
	else
	{
		SetFocusAsyncAction* pSetFocusAsync = new SetFocusAsyncAction(pTargetElement);

		// Put the action object into the queue and post a message.
		m_asyncActionsQueue.push_back(pSetFocusAsync);
		this->PostMessage(WM_APP_ASYNC_ACTION);

		return S_OK;
	}
}


STDMETHODIMP CExplorerPlugin::SetFocusAwayFromElement(IHTMLElement* pTargetElement, BOOL bGenerateOnChange, VARIANT_BOOL vbAsync)
{
	if (VARIANT_FALSE == vbAsync)
	{
		SetFocusAwayAsyncAction setFocusAwayAction(pTargetElement, bGenerateOnChange);
		return setFocusAwayAction.DoAction();
	}
	else
	{
		SetFocusAwayAsyncAction* pSetFocusAwayAsync = new SetFocusAwayAsyncAction(pTargetElement, bGenerateOnChange);

		// Put the action object into the queue and post a message.
		m_asyncActionsQueue.push_back(pSetFocusAwayAsync);
		this->PostMessage(WM_APP_ASYNC_ACTION);

		return S_OK;
	}
}


STDMETHODIMP CExplorerPlugin::FireEventOnElement(IHTMLElement* pTargetElement, BSTR bstrEventName, LONG wCharToRise, VARIANT_BOOL vbAsync)
{
	if (VARIANT_FALSE == vbAsync)
	{
		FireEventAsyncAction fireEventAction(pTargetElement, bstrEventName, wCharToRise);
		return fireEventAction.DoAction();
	}
	else
	{
		FireEventAsyncAction* pFireEventAsync = new FireEventAsyncAction(pTargetElement, bstrEventName, wCharToRise);

		// Put the action object into the queue and post a message.
		m_asyncActionsQueue.push_back(pFireEventAsync);
		this->PostMessage(WM_APP_ASYNC_ACTION);

		return S_OK;
	}
}


STDMETHODIMP CExplorerPlugin::ClickElementAt(IHTMLElement* pElement, LONG x, LONG y, LONG nRightClick)
{
	return HtmlHelpers::ClickElementAt(pElement, x, y, nRightClick);
}


STDMETHODIMP CExplorerPlugin::ClickElementAtAsync(IHTMLElement* pElement, LONG x, LONG y, LONG nRightClick)
{
	ClickAsyncAction* pClickAsync = new ClickAsyncAction(pElement, x, y, nRightClick);

	// Put the action object into the queue and post a message.
	m_asyncActionsQueue.push_back(pClickAsync);
	this->PostMessage(WM_APP_ASYNC_ACTION);

	return S_OK;
}


LRESULT CExplorerPlugin::OnAsyncAction(UINT uMsg, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	ATLASSERT(WM_APP_ASYNC_ACTION == uMsg);

	if (!m_asyncActionsQueue.empty())
	{
		IAsyncAction* pAsyncAction = m_asyncActionsQueue.front();

		pAsyncAction->DoAction();

		// Delete the object and remove it from the queue.
		delete pAsyncAction;
		pAsyncAction = NULL;
		m_asyncActionsQueue.pop_front();
	}
	else
	{
		traceLog << "m_asyncActionsQueue is empty in CExplorerPlugin::OnAsyncAction\n";
		ATLASSERT(FALSE && _T("m_asyncActionsQueue is empty in CExplorerPlugin::OnAsyncAction"));
	}

	return 0;
}


// Returns HRES_OPERATION_NOT_APPLICABLE if applied to a non combo box HTML control.
STDMETHODIMP CExplorerPlugin::FindSelectedOption(IHTMLElement* pElement, IHTMLElement** ppSelectedOption)
{
	if ((NULL == pElement) || (NULL == ppSelectedOption))
	{
		traceLog << "Invalid input in CExplorerPlugin::FindSelectedOption\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLSelectElement> spSelect = pElement;
	if (spSelect == NULL)
	{
		traceLog << "pElement is not a <select> element in CExplorerPlugin::FindSelectedOption\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}

	// Just let selectedOption property to be called for multiple selection lists (to accomodate list type changes).
	/*if (!HtmlHelpers::IsComboControl(spSelect) && HtmlHelpers::IsMultipleListControl(spSelect))
	{
		traceLog << "pElement is not a combo-box HTML control in CExplorerPlugin::FindSelectedOption\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}*/

	HRESULT hRes = S_OK;
	try
	{
		CComQIPtr<IHTMLElement> spResult = FindElementAlgorithms::FindSelectedOption(spSelect);
		if (spResult == NULL)
		{
			hRes = S_FALSE;
		}

		*ppSelectedOption = spResult.Detach();
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return hRes;
}


STDMETHODIMP CExplorerPlugin::FindAllSelectedOptions(IHTMLElement* pElement, IUnknown** ppResult)
{
	if ((NULL == pElement) || (NULL == ppResult))
	{
		traceLog << "Invalid input in CExplorerPlugin::FindAllSelectedOptions\n";
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLSelectElement> spSelect = pElement;
	if (spSelect == NULL)
	{
		traceLog << "pElement is not a <select> element in CExplorerPlugin::FindAllSelectedOptions\n";
		return HRES_OPERATION_NOT_APPLICABLE;
	}

	HRESULT hRes = S_OK;
	try
	{
		CComQIPtr<IUnknown, &IID_IUnknown> spResult = FindElementAlgorithms::FindAllSelectedOptions(spSelect);
		CComQIPtr<ILocalElementCollection> spColl   = spResult;

		if (spColl == NULL)
		{
			return E_FAIL;
		}

		// If the result collection is empty re-try in searchTimeout.
		BOOL bIsEmpty;
		if (SUCCEEDED(IsEmptyCollection(spColl, bIsEmpty)))
		{
			if (bIsEmpty)
			{
				hRes = S_FALSE;
			}
		}
		else
		{
			return E_FAIL;
		}

		*ppResult = spResult.Detach();
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		return E_FAIL;
	}

	return hRes;
}


HRESULT CExplorerPlugin::IsEmptyCollection(CComQIPtr<ILocalElementCollection> spColl, BOOL& bIsEmpty)
{
	LONG    nElemNumber = 0;
	LONG    nPageNumber = 0;
	LONG    nPageSize   = 0;

	HRESULT hRes = spColl->GetCollectionInfo(&nElemNumber, &nPageNumber, &nPageSize);
	if (FAILED(hRes))
	{
		bIsEmpty = TRUE;
		return hRes;
	}

	bIsEmpty = (0 == nElemNumber);
	return S_OK;
}



// IAccessible.
STDMETHODIMP CExplorerPlugin::get_accParent(IDispatch **ppdispParent)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accChildCount(long *pcountChildren)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accChild(VARIANT varChild, IDispatch **ppdispChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accName(VARIANT varChild, BSTR *pszName)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accDescription(VARIANT varChild, BSTR *pszDescription)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accHelp(VARIANT varChild, BSTR *pszHelp)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accHelpTopic(BSTR *pszHelpFile, VARIANT varChild, long *pidTopic)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accFocus(VARIANT *pvarChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accSelection(VARIANT *pvarChildren)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::get_accDefaultAction(VARIANT varChild, BSTR *pszDefaultAction)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::accSelect(long flagsSelect, VARIANT varChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::accLocation(long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::accNavigate(long navDir, VARIANT varStart, VARIANT *pvarEndUpAt)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::accHitTest(long xLeft, long yTop, VARIANT *pvarChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::accDoDefaultAction(VARIANT varChild)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::put_accName(VARIANT varChild, BSTR szName)
{
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerPlugin::put_accValue(VARIANT varChild, BSTR szValue)
{
	return E_NOTIMPL;
}


LRESULT CExplorerPlugin::OnAccGetObject(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
	return LresultFromObject(IID_IAccessible, wParam, GetUnknown());
}


STDMETHODIMP CExplorerPlugin::PostInputText(LONG nIEWnd, BSTR bstrText, BOOL bIsInputFile)
{
	HWND hIeWnd = (HWND)LongToHandle(nIEWnd);
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));
	ATLASSERT(bstrText != NULL);

	BOOL bRes = Keyboard::InputText(bstrText, hIeWnd, !bIsInputFile);
	return bRes ? S_OK : E_FAIL;
}


void CExplorerPlugin::CleanAyncActionsQueue()
{
	for (std::list<IAsyncAction*>::iterator it = m_asyncActionsQueue.begin();
	     it != m_asyncActionsQueue.end(); ++it)
	{
		delete (*it);
		*it = NULL;
	}

	m_asyncActionsQueue.clear();
}


STDMETHODIMP CExplorerPlugin::SetForceLoaded(LONG nIeWnd)
{
	// In Twebst Automation Studio we attach to browser on form onload and "Internet Explorer_Server" does not exists yet.
	HWND hIeWnd = static_cast<HWND>(LongToHandle(nIeWnd));
	if ((Common::GetWndClass(hIeWnd) != _T("Internet Explorer_Server")) && 
		(::GetCurrentThreadId() != ::GetWindowThreadProcessId(hIeWnd, NULL)))
	{
		return E_FAIL;
	}

	m_bIsForceLoaded = TRUE;
	m_hIeWnd         = hIeWnd;

	return S_OK;
}


CComQIPtr<IHTMLDocument2> CExplorerPlugin::FindDocumentFromPoint(LONG x, LONG y)
{
	if (Common::GetWndClass(m_hIeWnd) != _T("Internet Explorer_Server"))
	{
		return CComQIPtr<IHTMLDocument2>();
	}

	CComQIPtr<IAccessible> spWndAcc;
	HRESULT hRes = ::AccessibleObjectFromWindow(m_hIeWnd, OBJID_WINDOW, IID_IAccessible, (void**)&spWndAcc);
	if (FAILED(hRes) || !spWndAcc)
	{
		traceLog << "AccessibleObjectFromWindow failed in CExplorerPlugin::FindDocumentFromPoint\n";
		return CComQIPtr<IHTMLDocument2>();
	}

	CComQIPtr<IAccessible> spLastPane;
	CComQIPtr<IAccessible> spCrntAcc = spWndAcc;

	while (TRUE)
	{
		// Get the role of the current AA object.
		CComVariant vRole;
		hRes = spCrntAcc->get_accRole(CComVariant(CHILDID_SELF), &vRole);

		// Keep the last pane object.
		if (SUCCEEDED(hRes) && (VT_I4 == vRole.vt) && (ROLE_SYSTEM_CLIENT == vRole.lVal))
		{
			spLastPane = spCrntAcc;
		}

		CComVariant vAccCrntChild;
		hRes = spCrntAcc->accHitTest(x, y, &vAccCrntChild);

		if (FAILED(hRes) || (VT_EMPTY == vAccCrntChild.vt))
		{
			traceLog << "accHitTest failed in CExplorerPlugin::FindDocumentFromPoint\n";
			return CComQIPtr<IHTMLDocument2>();
		}

		if ((VT_DISPATCH == vAccCrntChild.vt) && (vAccCrntChild.pdispVal != NULL))
		{
			spCrntAcc = vAccCrntChild.pdispVal;
		}
		else
		{
			break;
		}
	}

	if (!spLastPane)
	{
		traceLog << "spLastPane is NULL in CExplorerPlugin::FindDocumentFromPoint\n";
		return CComQIPtr<IHTMLDocument2>();
	}

	CComQIPtr<IHTMLWindow2> spHtmlWindow = HtmlHelpers::AccessibleToHtmlWindow(spLastPane);
	if (!spHtmlWindow)
	{
		traceLog << "HtmlHelpers::AccessibleToHtmlWindow failed in CExplorerPlugin::FindDocumentFromPoint\n";
		return CComQIPtr<IHTMLDocument2>();
	}

	CComQIPtr<IHTMLDocument2> spDoc = HtmlHelpers::HtmlWindowToHtmlDocument(spHtmlWindow);
	if (!spDoc)
	{
		traceLog << "Cannot get document in CExplorerPlugin::FindDocumentFromPoint\n";
	}

	return spDoc;
}


STDMETHODIMP CExplorerPlugin::FindElementFromPoint(LONG x, LONG y, IHTMLElement** ppElem)
{
	if (!ppElem)
	{
		return E_INVALIDARG;
	}

	CComQIPtr<IHTMLDocument2> spDoc = FindDocumentFromPoint(x, y);
	if (!spDoc)
	{
		traceLog << "Cannot get spDoc in CExplorerPlugin::FindElementFromPoint\n";
		return S_OK;
	}

	// Get AA object coresponding to HTML document.
	CComQIPtr<IAccessible> spDocAcc = HtmlHelpers::HtmlDocumentToAccessible(spDoc);
	if (!spDocAcc)
	{
		traceLog << "Cannot get spDocAcc in CExplorerPlugin::FindElementFromPoint\n";
		return S_OK;
	}

	// Get AA document screen rectangle.
	long nScrDocLeft = 0;
	long nScrDocTop  = 0;
	long nDocWidth   = 0;
	long nDocHeight  = 0;

	HRESULT hRes = spDocAcc->accLocation(&nScrDocLeft, &nScrDocTop, &nDocWidth, &nDocHeight, CComVariant(CHILDID_SELF));
	if (FAILED(hRes))
	{
		traceLog << "spDocAcc->accLocation failed in CExplorerPlugin::FindElementFromPoint\n";
		return S_OK;
	}

	// Convert x, y in doc coordinates.
	x = x - nScrDocLeft;
	y = y - nScrDocTop;

	// Find zoom-level.
	int nZoomLevel = GetZoomLevel();

	// Scale x, y with zoom level.
	x = (LONG)((x * 100.0) / nZoomLevel);
	y = (LONG)((y * 100.0) / nZoomLevel);

	// Get element from point.
	CComQIPtr<IHTMLElement> spResElem;
	hRes = spDoc->elementFromPoint(x, y, &spResElem);

	if (SUCCEEDED(hRes) && spResElem)
	{
		*ppElem = spResElem.Detach();
	}
	else
	{
		traceLog << "spDoc->elementFromPoint failed in CExplorerPlugin::FindElementFromPoint\n";
	}

	return S_OK;
}


int CExplorerPlugin::GetZoomLevel()
{
	// Get the IWebBrowser2 pointer.
	int                     nZoomLevel = 100;
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT                 hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);

	if (SUCCEEDED(hRes) && spBrws)
	{
		// It seems the call below does NOT work from outside IE.
		CComVariant vZoom;
		hRes = spBrws->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, NULL, &vZoom);

		if (SUCCEEDED(hRes))
		{
			nZoomLevel = vZoom.lVal;
		}
	}

	return nZoomLevel;
}
