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
#include "Browser.h"
#include "Core.h"
#include "HtmlHelpIDs.h"
#include "FindInTimeout.h"
#include "Frame.h"
#include "SearchContext.h"
#include "..\OTWBSTInjector\OTWBSTInjector.h"
#include "MethodAndPropertyNames.h"
#include "HtmlHelpers.h"
#include "SearchCondition.h"
#include "MarshalService.h"
using namespace Common;


const ULONG BRWS_SHOW_IN_FOREGROUND = 0x01;
const ULONG BRWS_SHOW_MINIMIZE      = 0x02;
const ULONG BRWS_SHOW_MAXIMIZE      = 0x03;
const ULONG BRWS_SHOW_HIDE          = 0x04;
const ULONG BRWS_SHOW_VISIBLE       = 0x05;
const ULONG BRWS_SHOW_RESTORE       = 0x06;


STDMETHODIMP CBrowser::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IBrowser
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}

	return S_FALSE;
}


STDMETHODIMP CBrowser::get_nativeBrowser(IWebBrowser2** pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CBrowser::get_nativeBrowser\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_BROWSER_PROPERTY, IDH_BROWSER_NATIVE_BROWSER);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, NATIVE_BROWSER_PROPERTY, IDH_BROWSER_NATIVE_BROWSER))
	{
		traceLog << "Invalid state in CBrowser::get_nativeBrowser\n";
		return HRES_FAIL;
	}

	IWebBrowser2* pWebBrws = NULL;
	HRESULT hRes = m_spPlugin->GetNativeBrowser(&pWebBrws);

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		ATLASSERT(pWebBrws == NULL);
		traceLog << "Connection with the browser lost while calling m_spPlugin->GetNativeBrowser in CBrowser::get_nativeBrowser\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_NATIVE_BROWSER);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (pWebBrws != NULL)
	{
		*pVal = pWebBrws;
		return HRES_OK;
	}
	else
	{
		ATLASSERT(!SUCCEEDED(hRes));

		traceLog << "m_spPlugin->GetNativeBrowser failed with code " << hRes << "in CBrowser::get_nativeBrowser\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NATIVE_BROWSER_PROPERTY, IDH_BROWSER_NATIVE_BROWSER);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
}


HRESULT CBrowser::GetWebBrowser(IWebBrowser2** ppWebBrowser)
{
	if (NULL == ppWebBrowser)
	{
		traceLog << "ppWebBrowser parameter is NULL in Browser::GetWebBrowser\n";
		return E_INVALIDARG;
	}

	if (m_spPlugin == NULL)
	{
		traceLog << "m_spPlugin is NULL in CBrowser::GetWebBrowser\n";
		return E_UNEXPECTED;
	}

	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = m_spPlugin->GetNativeBrowser(&spBrws);
	if (spBrws == NULL)
	{
		traceLog << "m_spPlugin->GetNativeBrowser failed with code " << hRes << "in CBrowser::GetWebBrowser\n";
		return hRes;
	}

	*ppWebBrowser = spBrws.Detach();
	return S_OK;
}


STDMETHODIMP CBrowser::get_title(BSTR* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CBrowser::get_title\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, TITLE_PROPERTY, IDH_BROWSER_TITLE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, TITLE_PROPERTY, IDH_BROWSER_TITLE))
	{
		traceLog << "Invalid state in CBrowser::get_title\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->GetBrowserTitle(pVal);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling GetBrowserTitle in CBrowser::get_title\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_TITLE);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}
	else if (FAILED(hRes))
	{
		traceLog << "GetBrowserTitle failed in CBrowser::get_title with code:" << hRes << "\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, TITLE_PROPERTY, IDH_BROWSER_TITLE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CBrowser::get_app(BSTR* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CBrowser::get_app\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, APP_PROPERTY, IDH_BROWSER_APP);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, APP_PROPERTY, IDH_BROWSER_APP))
	{
		traceLog << "Invalid state in CBrowser::get_app\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->GetAppName(pVal);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling m_spPlugin->IsLoading in CBrowser::get_app\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_APP);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->GetAppName failed in CBrowser::get_app\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, APP_PROPERTY, IDH_BROWSER_APP);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CBrowser::get_url(BSTR* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CBrowser::get_url\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, URL_PROPERTY, IDH_BROWSER_URL);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetWebBrowser(&spBrws);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		ATLASSERT(spBrws == NULL);
		traceLog << "Connection with the browser lost while calling GetWebBrowser in CBrowser::get_url\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_URL);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (spBrws == NULL)
	{
		ATLASSERT(FAILED(hRes));
		traceLog << "Can not get the IWebBrowser2 object in CBrowser::get_url\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, URL_PROPERTY, IDH_BROWSER_URL);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	hRes = spBrws->get_LocationURL(pVal);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling spBrws->get_LocationURL in CBrowser::get_url\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_URL);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "IWebBrowser2::get_LocationURL failed in CBrowser::get_url with code:" << hRes << "\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, URL_PROPERTY, IDH_BROWSER_URL);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CBrowser::get_isLoading(VARIANT_BOOL* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal parameter is NULL in CBrowser::get_isLoading\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, IS_LOADING_PROPERTY, IDH_BROWSER_ISLOADING);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, IS_LOADING_PROPERTY, IDH_BROWSER_ISLOADING))
	{
		traceLog << "Invalid state in CBrowser::get_isLoading\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->IsLoading(pVal);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling m_spPlugin->IsLoading in CBrowser::get_isLoading\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_ISLOADING);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->IsLoading failed in CBrowser::get_isLoading\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, IS_LOADING_PROPERTY, IDH_BROWSER_ISLOADING);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


// *pVal is VARIANT_TRUE if the page is successfuly loaded, VARIANT_FALSE otherwise.
// If the page is not successfuly loaded:
// {
//   - if Core::get_loadTimeoutIsError is true return HRES_LOAD_TIMEOUT_ERR
//   - if Core::get_loadTimeoutIsError is false return HRES_LOAD_TIMEOUT_WARN
// }
// If Core::get_loadTimeout is zero check only the browser against the safe array condition.
STDMETHODIMP CBrowser::WaitToLoad(BSTR bstrCond, VARIANT_BOOL* pVal)
{
	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	FIRE_CANCEL_REQUEST();
	this->Sleep(Common::INTERNAL_GLOBAL_PAUSE * 5);

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	std::list<DescriptorToken> tokenList;
	int  nSafeArraySize;
	BOOL bInvalidArg = FALSE;

	if ((NULL == pVal) || (NULL == pVarArgs) || !SafeArraySize(pVarArgs, &nSafeArraySize))
	{
		bInvalidArg = TRUE;
	}
	else
	{
		BOOL bValidArgs = Common::GetDescriptorTokensList(pVarArgs, &tokenList);
		if (!bValidArgs || !IsValidDescriptorList(tokenList))
		{
			bInvalidArg = TRUE;
		}
	}

	if (bInvalidArg)
	{
		traceLog << "pVal or pVarArgs parameter is NULL or can not get the safe array size or invalid attributes in search condition in CBrowser::WaitToLoad\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, WAIT_TO_LOAD_METHOD, IDH_BROWSER_WAIT_TO_LOAD);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, WAIT_TO_LOAD_METHOD, IDH_BROWSER_WAIT_TO_LOAD))
	{
		traceLog << "Invalid state in CBrowser::WaitToLoad\n";
		return HRES_FAIL;
	}

	// Get the loadTimeout Core property value.
	LONG    nLoadTimeout = 0;
	HRESULT hRes         = m_spCore->get_loadTimeout(&nLoadTimeout);

	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeout failed in CBrowser::WaitToLoad with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, WAIT_TO_LOAD_METHOD, IDH_BROWSER_WAIT_TO_LOAD);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the loadTimeoutIsError property.
	VARIANT_BOOL vbLoadTimeoutIsErr;
	hRes = m_spCore->get_loadTimeoutIsError(&vbLoadTimeoutIsErr);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeoutIsError failed in CBrowser::WaitToLoad with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, WAIT_TO_LOAD_METHOD, IDH_BROWSER_WAIT_TO_LOAD);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	*pVal = VARIANT_FALSE;

	DWORD dwStartTime = ::GetTickCount();
	while (TRUE)
	{
		FIRE_CANCEL_REQUEST();

		VARIANT_BOOL vbIsLoading;
		HRESULT hRes = m_spPlugin->IsLoading(&vbIsLoading);
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost while calling m_spPlugin->IsLoading in CBrowser::WaitToLoad\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_WAIT_TO_LOAD);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}

		if (FAILED(hRes))
		{
			// Continue waiting in timeout.
			traceLog << "m_spPlugin->IsLoading failed in CBrowser::WaitToLoad\n";
		}
		else if (VARIANT_FALSE == vbIsLoading)
		{
			*pVal = VARIANT_TRUE;
			break;
		}

		DWORD dwCurrentTime = ::GetTickCount();
		if ((dwCurrentTime - dwStartTime) > (nLoadTimeout * TIME_SCALE))
		{
			// Timeout has expired. Quit the loop.
			break;
		}

		// Sleep for a while.
		this->Sleep(Common::INTERNAL_GLOBAL_PAUSE, TRUE);
	}

	if (0 == nLoadTimeout)
	{
		// If Core.loadTimeout is zero then don't wait the browser to load.
		*pVal = VARIANT_TRUE;
	}

	if (*pVal == VARIANT_FALSE)
	{
		// The page was NOT successfuly loaded.
		if (VARIANT_TRUE == vbLoadTimeoutIsErr)
		{
			// The script will eventually throw an exception.
			traceLog << "Page load timeout in CBrowser::WaitToLoad \n";
			SetComErrorMessage(IDS_WAIT_TO_LOAD_TIMEOUT, IDH_BROWSER_WAIT_TO_LOAD);
			SetLastErrorCode(ERR_LOAD_TIMEOUT);
			return HRES_LOAD_TIMEOUT_ERR;
		}
		else
		{
			// No exception is thrown by the script.
			ATLASSERT(VARIANT_FALSE == vbLoadTimeoutIsErr);
			SetLastErrorCode(ERR_LOAD_TIMEOUT);
			return HRES_LOAD_TIMEOUT_WARN;
		}
	}
	else
	{
		// The page was successfuly loaded.
		ATLASSERT(VARIANT_TRUE == *pVal);

		if (0 != nSafeArraySize)
		{
			LONG nSearchFlags = 0;
			hRes = m_spPlugin->CheckBrowserDescriptor(nSearchFlags, pVarArgs);
			if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
			{
				traceLog << "Connection with the browser lost while calling m_spPlugin->CheckBrowserDescriptor in CBrowser::WaitToLoad\n";
				SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_WAIT_TO_LOAD);
				SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
				return HRES_BRWS_CONNECTION_LOST_ERR;
			}

			if (S_FALSE == hRes)
			{
				// The browser does not match the safe array.
				*pVal = VARIANT_FALSE;
			}
			else if (hRes != S_OK)
			{
				traceLog << "m_spPlugin->CheckBrowserDescriptor failed in CBrowser::WaitToLoad\n";
				SetComErrorMessage(IDS_METHOD_CALL_FAILED, WAIT_TO_LOAD_METHOD, IDH_BROWSER_WAIT_TO_LOAD);
				SetLastErrorCode(ERR_FAIL);
				return HRES_FAIL;
			}
		}

		return HRES_OK;
	}
}


STDMETHODIMP CBrowser::Navigate(BSTR bstrUrl)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (bstrUrl == NULL)
	{
		traceLog << "pVal parameter is NULL in CBrowser::Navigate\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, NAVIGATE_METHOD, IDH_BROWSER_NAVIGATE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetWebBrowser(&spBrws);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		ATLASSERT(spBrws == NULL);
		traceLog << "Connection with the browser lost while calling GetWebBrowser in CBrowser::Navigate\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_NAVIGATE);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (spBrws == NULL)
	{
		ATLASSERT(FAILED(hRes));
		traceLog << "Can not get the IWebBrowser2 object in CBrowser::Navigate\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, NAVIGATE_METHOD, IDH_BROWSER_NAVIGATE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	CComVariant vDummy;
	hRes = spBrws->Navigate(bstrUrl, &vDummy, &vDummy, &vDummy, &vDummy);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling spBrws->Navigate in CBrowser::Navigate\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_NAVIGATE);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "IWebBrowser2::Navigate failed in CBrowser::Navigate with code:" << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, NAVIGATE_METHOD, IDH_BROWSER_NAVIGATE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Give Internet Explorer browser a chance to start the navigation.
	this->Sleep(Common::INTERNAL_GLOBAL_PAUSE, TRUE);
	return HRES_OK;
}


STDMETHODIMP CBrowser::Close(void)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetWebBrowser(&spBrws);
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		ATLASSERT(spBrws == NULL);
		traceLog << "Connection with the browser lost while calling GetWebBrowser in CBrowser::Close\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_CLOSE);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (spBrws == NULL)
	{
		ATLASSERT(FAILED(hRes));
		traceLog << "Can not get the IWebBrowser2 object in CBrowser::Close, code=" << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, CLOSE_METHODE, IDH_BROWSER_CLOSE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	hRes = spBrws->Quit();
	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling spBrws->Quit in CBrowser::Close\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_CLOSE);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}

	if (FAILED(hRes))
	{
		traceLog << "IWebBrowser2::Quit failed in CBrowser::Close with code:" << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, CLOSE_METHODE, IDH_BROWSER_CLOSE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return hRes;
}


STDMETHODIMP CBrowser::get_core(ICore** pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CBrowser::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in CBrowser::get_core\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, CORE_PROPERTY, IDH_BROWSER_CORE);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	CComQIPtr<ICore> spCore = m_spCore;
	*pVal = spCore.Detach();

	return HRES_OK;
}


STDMETHODIMP CBrowser::get_topFrame(IFrame** ppTopFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppTopFrame)
	{
		traceLog << "ppTopFrame is NULL in CBrowser::get_topFrame\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, TOP_FRAME_PROPERTY, IDH_BORWSER_TOP_FRAME);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, TOP_FRAME_PROPERTY, IDH_BORWSER_TOP_FRAME))
	{
		traceLog << "Invalid state in CBrowser::get_topFrame\n";
		return HRES_FAIL;
	}

	CComQIPtr<IHTMLWindow2> spTopFrame;
	HRESULT hRes = m_spPlugin->GetTopFrame(&spTopFrame);
	if ((hRes != S_OK) || (spTopFrame == NULL))
	{
		traceLog << "m_spPlugin->GetTopFrame failed in CBrowser::get_topFrame with code:" << hRes <<"\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, TOP_FRAME_PROPERTY, IDH_BORWSER_TOP_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Create an IFrame object.
	hRes = CreateFrameObject(spTopFrame, ppTopFrame);
	if (NULL == *ppTopFrame)
	{
		ATLASSERT(FAILED(hRes));

		traceLog << "Can not create a Frame object in CBrowser::get_topFrame\n";
		SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_BORWSER_TOP_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


STDMETHODIMP CBrowser::FindElement(BSTR bstrTag, BSTR bstrCond, IElement** ppElement)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	if (NULL == ppElement)
	{
		traceLog << "ppElement is null in CBrowser::FindElement\n";
		SetComErrorMessage(IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDH_BROWSER_FIND_ELEMENT);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY;
	SearchContext sc("CBrowser::FindElement", nSearchFlags, IDH_BROWSER_FIND_ELEMENT,
	                 IDS_ERR_FIND_ELEMENT_INVALID_PARAM, IDS_ERR_FIND_ELEMENT_FAILED,
					 IDS_FIND_ELEMENT_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(NULL, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElementObject(spResult, ppElement);
		if (NULL == *ppElement)
		{
			traceLog << "Can not create an Element object in CBrowser::FindElement\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT, IDH_BROWSER_FIND_ELEMENT);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CBrowser::FindFrame(BSTR bstrCond, IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	if (NULL == ppFrame)
	{
		traceLog << "ppFrame is null in CBrowser::FindFrame\n";
		SetComErrorMessage(IDS_ERR_FIND_FRAME_INVALID_PARAM, IDH_BROWSER_FIND_FRAME);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_FRAME | SEARCH_ALL_HIERARCHY;
	SearchContext sc("CBrowser::FindFrame", nSearchFlags, IDH_BROWSER_FIND_FRAME,
	                 IDS_ERR_FIND_FRAME_INVALID_PARAM, IDS_ERR_FIND_FRAME_FAILED,
                     IDS_FIND_FRAME_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(NULL, sc, NULL, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateFrameObject(spResult, ppFrame);
		if (NULL == *ppFrame)
		{
			traceLog << "Can not create a Frame object in CBrowser::FindFrame\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_BROWSER_FIND_FRAME);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CBrowser::FindAllElements(BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	if (NULL == ppElementList)
	{
		traceLog << "ppElementList is null in CBrowser::FindAllElements\n";
		SetComErrorMessage(IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDH_BROWSER_FIND_ALL_ELEMENTS);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;	
	}

	ULONG nSearchFlags = SEARCH_ELEMENT | SEARCH_ALL_HIERARCHY | SEARCH_COLLECTION;
	SearchContext sc("CBrowser::FindAllElements", nSearchFlags, IDH_BROWSER_FIND_ALL_ELEMENTS,
	                 IDS_ERR_FIND_ALL_ELEMENTS_INVALID_PARAM, IDS_ERR_FIND_ALL_ELEMENTS_FAILED,
                     IDS_FIND_ALL_ELEMENTS_LOAD_TIMEOUT);
	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(NULL, sc, bstrTag, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateElemListObject(spResult, ppElementList);
		if (NULL == *ppElementList)
		{
			traceLog << "Can not create an ElementList object in CBrowser::FindAllElements\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_ELEMENT_LIST, IDH_BROWSER_FIND_ALL_ELEMENTS);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CBrowser::get_navigationError(LONG* pVal)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == pVal)
	{
		traceLog << "pVal is NULL in CBrowser::get_navigationError\n";
		SetComErrorMessage(IDS_PROPERTY_FAILED, NAVIGATION_ERROR, IDH_BROWSER_NAVIGATION_ERR);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_PROPERTY_FAILED, NAVIGATION_ERROR, IDH_BROWSER_NAVIGATION_ERR))
	{
		traceLog << "Invalid state in CBrowser::get_navigationError\n";
		return HRES_FAIL;
	}

	HRESULT hRes = m_spPlugin->GetLastNavigationErr(pVal);
	if (FAILED(hRes))
	{
		if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
		{
			traceLog << "Connection with the browser lost in CBrowser::get_navigationError\n";
			SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, IDH_BROWSER_NAVIGATION_ERR);
			SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
			return HRES_BRWS_CONNECTION_LOST_ERR;
		}
		else
		{
			traceLog << "m_spPlugin->GetLastNavigationErr failed with code " << hRes << " in CBrowser::get_navigationError\n";
			SetComErrorMessage(IDS_PROPERTY_FAILED, NAVIGATION_ERROR, IDH_BROWSER_NAVIGATION_ERR);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return HRES_OK;
}


BOOL CBrowser::IsValidDescriptorList(const std::list<DescriptorToken>& tokens)
{
	for (std::list<DescriptorToken>::const_iterator it = tokens.begin();
	     it != tokens.end(); ++it)
	{
		const String& sAttributeName = it->m_sAttributeName;
		if (_tcsicmp(sAttributeName.c_str(), BROWSER_TITLE_ATTR_NAME) &&
		    _tcsicmp(sAttributeName.c_str(), BROWSER_URL_ATTR_NAME)   &&
			_tcsicmp(sAttributeName.c_str(), BROWSER_PID_ATTR_NAME)   &&
			_tcsicmp(sAttributeName.c_str(), BROWSER_TID_ATTR_NAME))
		{
			return FALSE;
		}
	}
	
	return TRUE;
}


STDMETHODIMP CBrowser::FindModelessHtmlDialog(BSTR bstrCond, IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	// The initial version of this method had a safe array input parameters for search conditions.
	// that's why the code below might seem strange and unoptimal.
	// To prevent regressions, bstrCond is converted to a safe array and the initial code remains unchanged.
	SearchCondition pVarArgs;
	pVarArgs.AddMultiCondition(bstrCond);

	if (NULL == ppFrame)
	{
		traceLog << "ppFrame is NULL in CBrowser::FindModelessHtmlDialog\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, BROWSER_FIND_HTML_DIALOG_METHOD, IDH_BROWSER_FIND_MODELESS_HTML_DLG);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	ULONG nSearchFlags = SEARCH_HTML_DIALOG;
	SearchContext sc("CBrowser::FindModelessHtmlDialog", nSearchFlags, IDH_BROWSER_FIND_MODELESS_HTML_DLG,
					 IDS_ERR_FIND_HTML_DIALOG_INVALID_PARAM, IDS_ERR_FIND_HTML_DIALOG_FAILED,
					 IDS_FIND_HTML_DIALOG_LOAD_TIMEOUT);

	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(NULL, sc, NULL, pVarArgs, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateFrameObject(spResult, ppFrame);
		if (NULL == *ppFrame)
		{
			traceLog << "Can not create a Frame object in CBrowser::FindModelessHtmlDialog\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_BROWSER_FIND_MODELESS_HTML_DLG);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


STDMETHODIMP CBrowser::FindModalHtmlDialog(IFrame** ppFrame)
{
	FIRE_CANCEL_REQUEST();

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (NULL == ppFrame)
	{
		traceLog << "ppFrame is NULL in CBrowser::FindModalHtmlDialog\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, BROWSER_FIND_MODAL_HTML_METHOD, IDH_BROWSER_FIND_MODAL_HTML_DLG);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	ULONG nSearchFlags = SEARCH_MODAL_HTML_DLG;
	SearchContext sc("CBrowser::FindModalHtmlDialog", nSearchFlags, IDH_BROWSER_FIND_MODAL_HTML_DLG,
					 IDS_ERR_FIND_MODAL_HTML_DIALOG_INVALID_PARAM, IDS_ERR_FIND_MODAL_HTML_DIALOG_FAILED,
					 IDS_FIND_MODAL_HTML_DIALOG_LOAD_TIMEOUT);

	CComQIPtr<IUnknown, &IID_IUnknown> spResult;
	HRESULT hRes = FindInContainer(NULL, sc, NULL, NULL, &spResult);

	if (spResult != NULL)
	{
		HRESULT hCreateElem = CreateFrameObject(spResult, ppFrame);
		if (NULL == *ppFrame)
		{
			traceLog << "Can not create a Frame object in CBrowser::FindModalHtmlDialog\n";
			SetComErrorMessage(IDS_ERR_CAN_NOT_CREATE_FRAME, IDH_BROWSER_FIND_MODAL_HTML_DLG);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	return hRes;
}


HRESULT CBrowser::ClosePopupOrPrompt(BSTR bstrPopupText, VARIANT vButton, BSTR bstrValue, BSTR* pPopupText)
{
	FIRE_CANCEL_REQUEST_NO_CLOSE_POPUP();

	LPCTSTR szMethodName = NULL;
	DWORD   dwHelpID     = 0;
	
	if (bstrValue != NULL)
	{
		szMethodName = BROWSER_CLOSE_PROMPT_METHOD;
		dwHelpID     = IDH_BROWSER_CLOSE_PROMPT;
	}
	else
	{
		szMethodName = BROWSER_CLOSE_POPUP_METHOD;
		dwHelpID     = IDH_BROWSER_CLOSE_POPUP;
	}

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((NULL == bstrPopupText) || (NULL == pPopupText) || !Common::IsValidOptionVariant(vButton))
	{
		traceLog << "Invalid input params in CBrowser::ClosePopupOrPrompt\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, szMethodName, dwHelpID))
	{
		traceLog << "Invalid state in CBrowser::ClosePopupOrPrompt\n";
		return HRES_FAIL;
	}

	LONG    nThreadID = 0;
	HRESULT hRes      = m_spPlugin->GetBrowserThreadID(&nThreadID);

	if (HRESULT_CODE(hRes) == RPC_S_SERVER_UNAVAILABLE)
	{
		traceLog << "Connection with the browser lost while calling m_spPlugin->GetBrowserThreadID in CBrowser::ClosePopupOrPrompt\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, dwHelpID);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}
	else if (FAILED(hRes))
	{
		traceLog << "m_spPlugin->GetBrowserThreadID failed in CBrowser::ClosePopupOrPrompt\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	String sText = bstrPopupText;

	// Get the searchTimeout property.
	LONG nSearchTimeout = 0;
	hRes = m_spCore->get_searchTimeout(&nSearchTimeout);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_searchTimeout failed for CBrowser::ClosePopupOrPrompt with code " << hRes << "\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	String sPopupText;
	HWND   hDialog     = NULL;
	DWORD  dwStartTime = ::GetTickCount();

	while (TRUE)
	{
		FIRE_CANCEL_REQUEST_NO_CLOSE_POPUP();

		if (!Common::GetPopupByText(nThreadID, sText, hDialog, &sPopupText))
		{
			traceLog << "Common::GetPopupByText failed in CBrowser::ClosePopupOrPrompt\n";
			SetComErrorMessage(IDS_METHOD_CALL_FAILED, szMethodName, dwHelpID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}

		if (hDialog != NULL)
		{
			break;
		}

		DWORD dwCurrentTime   = ::GetTickCount();
		DWORD dwElapsedTime   = dwCurrentTime - dwStartTime;

		if (dwElapsedTime > nSearchTimeout * Common::TIME_SCALE)
		{
			// Timeout.
			traceLog << "Search timeout for in CBrowser::ClosePopupOrPrompt\n";
			SetComErrorMessage(IDS_BROWSER_CLOSE_POPUP_TIMEOUT, szMethodName, dwHelpID);
			return HRES_FAIL;
		}

		// Sleep for a while.
		this->Sleep(Common::INTERNAL_GLOBAL_PAUSE);
	}

	ATLASSERT(hDialog != NULL);

	if (pPopupText != NULL)
	{
		*pPopupText = CComBSTR(sPopupText.c_str()).Detach();
	}

	if (bstrValue != NULL)
	{
		// It is a prompt; fill out the value field.
		HWND hEdit = Common::GetChildWindowByClassName(hDialog, _T("Edit"));
		if (NULL == hEdit)
		{
			traceLog << "Cannot find edit field in CBrowser::ClosePopupOrPrompt\n";
			SetComErrorMessage(IDS_CLOSE_POPUP_EDIT_NOT_FOUND, szMethodName, dwHelpID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}

		USES_CONVERSION;
		::SendMessage(hEdit, WM_SETTEXT, 0, (LPARAM)T2W(bstrValue));
	}

	BOOL bRes = FALSE;
	if (VT_BSTR == vButton.vt)
	{
		bRes = Common::PressButtonOnPopup(hDialog, vButton.bstrVal);
	}
	else
	{
		int nIndex;
		switch (vButton.vt)
		{
			case VT_I4:  nIndex = vButton.lVal;   break;
			case VT_I2:  nIndex = vButton.iVal;   break;
			case VT_UI4: nIndex = vButton.ulVal;  break;
			case VT_UI2: nIndex = vButton.uiVal;  break;
			case VT_INT: nIndex = vButton.intVal; break;
			default: ATLASSERT(FALSE);
		};

		// nIndex is zero based.
		bRes = Common::PressButtonOnPopup(hDialog, nIndex + 1);
	}

	if (!bRes)
	{
		traceLog << "Common::PressButtonOnPopup failed in CBrowser::ClosePopupOrPrompt\n";
		SetComErrorMessage(IDS_CLOSE_POPUP_BUTTON_NOT_FOUND, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
	else
	{
		return HRES_OK;
	}
}


STDMETHODIMP CBrowser::ClosePopup(BSTR bstrPopupText, VARIANT vButton, BSTR* pPopupText)
{
	return ClosePopupOrPrompt(bstrPopupText, vButton, NULL, pPopupText);
}


STDMETHODIMP CBrowser::ClosePrompt(BSTR bstrPromptText, BSTR bstrValue, VARIANT vButton, BSTR* pPopupText)
{
	return ClosePopupOrPrompt(bstrPromptText, vButton, bstrValue, pPopupText);
}


STDMETHODIMP CBrowser::GetAttr(BSTR bstrAttrName, VARIANT* pVal)
{
	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if (!bstrAttrName || !bstrAttrName[0] || !pVal)
	{
		traceLog << "Invalid arguments in CBrowser::GetAttr\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, BROWSER_GETATTR_METHOD, IDH_BROWSER_GET_ATTR);
		SetLastErrorCode(ERR_INVALID_ARG);

		return HRES_INVALID_ARG;
	}

	if (!IsValidState(IDS_METHOD_CALL_FAILED, BROWSER_GETATTR_METHOD, IDH_BROWSER_GET_ATTR))
	{
		traceLog << "Invalid state in CBrowser::GetAttr\n";
		return HRES_FAIL;
	}

	HRESULT hRes = HRES_OK;
	if (!_wcsicmp(bstrAttrName, L"appname"))
	{
		hRes = GetAppName(pVal);
	}
	else if (!_wcsicmp(bstrAttrName, L"pid"))
	{
		hRes = GetPID(pVal);
	}
	else
	{
		return E_INVALIDARG;
	}

	if (FAILED(hRes))
	{
		traceLog << "Failure in CBrowser::GetAttr\n";

		SetComErrorMessage(IDS_METHOD_CALL_FAILED, BROWSER_GETATTR_METHOD, IDH_BROWSER_GET_ATTR);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}
	else
	{
		return hRes;
	}
}


STDMETHODIMP CBrowser::SetAttr(BSTR bstrAttrName, VARIANT newVal)
{
	if (!bstrAttrName || !bstrAttrName[0])
	{
		traceLog << "Invalid arguments in CBrowser::SetAttr\n";
		return E_INVALIDARG;
	}

	return E_NOTIMPL;
}


HRESULT CBrowser::GetAppName(VARIANT* pAppName)
{
	ATLASSERT(pAppName);

	LONG nProcID = 0;
	HRESULT hRes = m_spPlugin->GetBrowserProcessID(&nProcID);
	if (FAILED(hRes) || !nProcID)
	{
		traceLog << "m_spPlugin->GetBrowserProcessID failed in CBrowser::GetAppName with code:" << hRes <<"\n";
		SetComErrorMessage(IDS_METHOD_CALL_FAILED, BROWSER_GETATTR_METHOD, IDH_BROWSER_GET_ATTR);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	CComVariant vPid(nProcID);
	return vPid.Detach(pAppName);
}


HRESULT CBrowser::GetPID(VARIANT* pPid)
{
	ATLASSERT(pPid);
	return E_NOTIMPL;
}


/////////////////////////////////////////////////////////////////////////////////////////
// Methods and properties not yet implemented.

STDMETHODIMP CBrowser::SaveSnapshot(BSTR bstrImgFileName)
{
	ATLASSERT(bstrImgFileName != NULL);
	ATLASSERT(FALSE && _T("Not yet implemented!"));

	return E_NOTIMPL;
}

STDMETHODIMP CBrowser::FindElementByXPath(BSTR bstrXPath, IElement** ppElement)
{
	ATLASSERT(bstrXPath != NULL);
	ATLASSERT(ppElement != NULL);

	return E_NOTIMPL;
}
