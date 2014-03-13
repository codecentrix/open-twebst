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
#include "HtmlHelpIds.h"
#include "DebugServices.h"
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "SafeArrayAutoPtr.h"
#include "CatCLSLib.h"
#include "CodeErrors.h"
#include "Registry.h"
#include "Settings.h"
#include "SearchContext.h"
#include "Core.h"
#include "Element.h"
#include "ElementList.h"
#include "Frame.h"
#include "BaseLibObject.h"


extern HINSTANCE g_hInstance;
BOOL BaseLibObject::s_bPopupsMapInit = FALSE;
CAtlMap<CString, int> BaseLibObject::s_autoClosePopupsMap;

static BOOL s_dummyResult = BaseLibObject::InitCloseBrowserPopups();


void BaseLibObject::SetPlugin(IExplorerPlugin* pPlugin)
{
	ATLASSERT(pPlugin != NULL);
	m_spPlugin = pPlugin;
}


void BaseLibObject::SetCore(ICore* pCore)
{
	ATLASSERT(pCore    != NULL);
	ATLASSERT(m_spCore == NULL);
	m_spCore = pCore;
}


void BaseLibObject::SetLastErrorCode(LONG nErr)
{
	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in BaseLibObject::SetLastErrorCode\n";
		return;
	}

	CCore* pCore = static_cast<CCore*>(m_spCore.p); // Down cast !!!
	pCore->SetLastErrorCode(nErr);
}


HRESULT BaseLibObject::SetComErrorMessage(UINT nId, DWORD dwHelpID)
{
	String sHelpFile;
	try
	{
		// Get the help file full path name.
		sHelpFile = Registry::RegGetStringValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                        Settings::REG_SETTINGS_INSTALLATION_DIRECTORY);
		if (_T('\\') != sHelpFile[sHelpFile.size() - 1])
		{
			sHelpFile += _T('\\');
		}

		USES_CONVERSION;
		sHelpFile += Settings::HELP_FILE_RELATIVE_PATH;
		return AtlReportError(m_clsid, nId, dwHelpID, T2W((LPTSTR)sHelpFile.c_str()));
	}
	catch (const RegistryException&)
	{
		// Can not get the help file full path name.
		return AtlReportError(m_clsid, nId);
	}
}


HRESULT BaseLibObject::SetComErrorMessage(UINT nIdPattern, LPCTSTR szFunctionName, DWORD dwHelpID)
{
	// Compute the error message.
	ATLASSERT(g_hInstance != NULL);
	TCHAR szPatternBuffer[128];
	int nRes = ::LoadString(g_hInstance, nIdPattern, szPatternBuffer, (sizeof(szPatternBuffer) / sizeof(szPatternBuffer[0])));
	ATLASSERT(nRes);

#pragma warning(disable:4996)
	TCHAR szErrorMsg[128];
	const int nBufferSize = sizeof(szErrorMsg) / sizeof(szErrorMsg[0]);
	_sntprintf(szErrorMsg, nBufferSize - 1, szPatternBuffer, szFunctionName);
	szErrorMsg[nBufferSize - 1] = _T('\0');
#pragma warning(default:4996)

	String sHelpFile;
	try
	{
		// Get the help file full path name.
		sHelpFile = Registry::RegGetStringValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME,
		                                        Settings::REG_SETTINGS_INSTALLATION_DIRECTORY);
		if (_T('\\') != sHelpFile[sHelpFile.size() - 1])
		{
			sHelpFile += _T('\\');
		}

		sHelpFile += Settings::HELP_FILE_RELATIVE_PATH;
		return AtlReportError(m_clsid, szErrorMsg, dwHelpID, sHelpFile.c_str());
	}
	catch (const RegistryException&)
	{
		// Can not get the help file full path name.
		return AtlReportError(m_clsid, szErrorMsg);
	}
}


HRESULT BaseLibObject::CreateElementObject(IUnknown* pHtmlElement, IElement** ppElement)
{
	ATLASSERT(pHtmlElement != NULL);
	ATLASSERT(ppElement    != NULL);
	ATLASSERT(m_spPlugin   != NULL);
	ATLASSERT(m_spCore     != NULL);

	// Query for IHTMLElement.
	CComQIPtr<IHTMLElement> spHtmlElement = pHtmlElement;
	if (spHtmlElement == NULL)
	{
		traceLog << "Can not query for IHTMLElement in BaseLibObject::CreateElementObject\n";
		return E_NOINTERFACE;
	}

	IElement* pNewElement = NULL;
	HRESULT   hRes = CComCoClass<CElement>::CreateInstance(&pNewElement);
	if (pNewElement != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CElement* pElementObject = static_cast<CElement*>(pNewElement); // Down cast !!!
		ATLASSERT(pElementObject != NULL);
		pElementObject->SetPlugin(m_spPlugin);
		pElementObject->SetCore(m_spCore);
		pElementObject->SetHtmlElement(spHtmlElement);

		// Set the object reference into the output pointer.
		*ppElement = pNewElement;
	}

	return hRes;
}


HRESULT BaseLibObject::CreateFrameObject(IUnknown* pHtmlFrame, IFrame** ppFrame)
{
	ATLASSERT(pHtmlFrame != NULL);
	ATLASSERT(ppFrame    != NULL);
	ATLASSERT(m_spPlugin != NULL);
	ATLASSERT(m_spCore   != NULL);

	// Query for IHTMLElement.
	CComQIPtr<IHTMLWindow2> spHtmlWindow = pHtmlFrame;
	if (spHtmlWindow == NULL)
	{
		traceLog << "Can not query for IHTMLWindow2 in BaseLibObject::CreateFrameObject\n";
		return E_NOINTERFACE;
	}

	IFrame* pNewFrame = NULL;
	HRESULT hRes = CComCoClass<CFrame>::CreateInstance(&pNewFrame);
	if (pNewFrame != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CFrame* pFrameObject = static_cast<CFrame*>(pNewFrame); // Down cast !!!
		ATLASSERT(pFrameObject != NULL);
		pFrameObject->SetPlugin(m_spPlugin);
		pFrameObject->SetCore(m_spCore);
		pFrameObject->SetHtmlWindow(spHtmlWindow);

		// Set the object reference into the output pointer.
		*ppFrame = pNewFrame;
	}

	return hRes;
}

HRESULT BaseLibObject::CreateElemListObject(IUnknown* pLocalElemCollection, IElementList** ppElementList)
{
	ATLASSERT(pLocalElemCollection != NULL);
	ATLASSERT(ppElementList        != NULL);
	ATLASSERT(m_spPlugin           != NULL);
	ATLASSERT(m_spCore             != NULL);

	// Query for IHTMLElement.
	CComQIPtr<ILocalElementCollection> spLocalElemCollection = pLocalElemCollection;
	if (spLocalElemCollection == NULL)
	{
		traceLog << "Can not query for ILocalElementCollection in BaseLibObject::CreateElementListObject\n";
		return E_NOINTERFACE;
	}

	IElementList* pNewElementList = NULL;
	HRESULT hRes = CComCoClass<CElementList>::CreateInstance(&pNewElementList);
	if (pNewElementList != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CElementList* pElemListObject = static_cast<CElementList*>(pNewElementList); // Down cast !!!
		ATLASSERT(pElemListObject != NULL);
		pElemListObject->SetPlugin(m_spPlugin);
		pElemListObject->SetCore(m_spCore);
		hRes = pElemListObject->InitElementList(spLocalElemCollection);

		if (FAILED(hRes))
		{
			traceLog << "ElementList::InitElementList failed with code " << hRes << " in BaseLibObject::CreateElementListObject\n";
			return hRes;
		}

		// Set the object reference into the output pointer.
		*ppElementList = pNewElementList;
	}

	return hRes;
}


HRESULT BaseLibObject::FindInContainer(IUnknown* pContainer, SearchContext& sc,
                                       BSTR bstrTag, SAFEARRAY* pVarArgs, IUnknown** ppResult)
{
	return Find(pContainer, sc, bstrTag, pVarArgs, (void**)ppResult);
}


HRESULT BaseLibObject::Find(IUnknown* pContainer, SearchContext& sc,
                            BSTR bstrTag, SAFEARRAY* pVarArgs, void** ppResult)
{
	ATLASSERT(NULL == *ppResult);

	// Reset the lastError property.
	SetLastErrorCode(ERR_OK);

	if ((NULL == bstrTag) &&
		 !(sc.m_nSearchContextFlags & Common::SEARCH_FRAME)           &&
		 !(sc.m_nSearchContextFlags & Common::SEARCH_HTML_DIALOG)     &&
		 !(sc.m_nSearchContextFlags & Common::SEARCH_SELECTED_OPTION) &&
		 !(sc.m_nSearchContextFlags & Common::SEARCH_MODAL_HTML_DLG))
	{
		traceLog << "bstrTag is NULL for " << sc.m_sFunctionName << "\n";
		SetComErrorMessage(sc.m_dwInvalidParamMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;	
	}

	SafeArrrayAutoPtr newSafeArray;
	int               nSafeArraySize = 0;

	// a NULL pVarArgs is accepted.
	if ((NULL != pVarArgs) && !SafeArraySize(pVarArgs, &nSafeArraySize))
	{
		traceLog << "Can not get the size of the safe array in BaseLibObject::Find\n";
		SetComErrorMessage(sc.m_dwInvalidParamMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (0 == nSafeArraySize)
	{
		// Create a new safe array.
		BOOL           bFail       = FALSE;
		SAFEARRAYBOUND rgsabound   = { 1, 0 };
		SAFEARRAY*     pNewVarArgs = ::SafeArrayCreate(VT_VARIANT, 1, &rgsabound);
		if (pNewVarArgs != NULL)
		{
			CComVariant varNew(L"src=*");
			long        nIndex = 0;
			HRESULT     hRes   = ::SafeArrayPutElement(pNewVarArgs, &nIndex, &varNew);
			if (SUCCEEDED(hRes))
			{
				pVarArgs = pNewVarArgs;
				newSafeArray.Attach(pNewVarArgs);	// To auto destroy on function exit.
			}
			else
			{
				bFail = TRUE;
			}
		}
		else
		{
			bFail = TRUE;
		}

		if (bFail)
		{
			traceLog << "Can not modify the safe array in BaseLibObject::Find\n";
			SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
			SetLastErrorCode(ERR_FAIL);
			return HRES_FAIL;
		}
	}

	if (NULL == ppResult)
	{
		traceLog << "ppResult is NULL for " << sc.m_sFunctionName << "\n";
		SetComErrorMessage(sc.m_dwInvalidParamMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL for " << sc.m_sFunctionName << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	if (m_spPlugin == NULL)
	{
		traceLog << "m_spPlugin is NULL for " << sc.m_sFunctionName << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the loadTimeout Core property value.
	LONG    nLoadTimeout = 0;
	HRESULT hRes         = m_spCore->get_loadTimeout(&nLoadTimeout);

	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeout failed for " << sc.m_sFunctionName << " with code " << hRes << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the loadTimeoutIsError property.
	VARIANT_BOOL vbLoadTimeoutIsErr;
	hRes = m_spCore->get_loadTimeoutIsError(&vbLoadTimeoutIsErr);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_loadTimeoutIsError failed for " << sc.m_sFunctionName << " with code " << hRes << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	// Get the searchTimeout property.
	LONG nSearchTimeout;
	hRes = m_spCore->get_searchTimeout(&nSearchTimeout);
	if (FAILED(hRes))
	{
		traceLog << "ICore::get_searchTimeout failed for " << sc.m_sFunctionName << " with code " << hRes << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	try
	{
		LONG nSearchFlags = sc.m_nSearchContextFlags;
		BOOL bLoadTimeout = FALSE;

		FindObjectInContainer finder(pContainer, m_spPlugin, nSearchFlags, bstrTag, pVarArgs, m_spCore);
		bLoadTimeout = FinderInTimeout::Find(&finder, nSearchTimeout, nLoadTimeout, vbLoadTimeoutIsErr, this, this);
		CComQIPtr<IUnknown, &IID_IUnknown> spResult = finder.GetResultObject();

		*ppResult = spResult.Detach();

		if (bLoadTimeout)
		{
			SetLastErrorCode(ERR_LOAD_TIMEOUT);
			if (VARIANT_TRUE == vbLoadTimeoutIsErr)
			{
				// The script will throw an exception.
				ATLASSERT(NULL == *ppResult);
				traceLog << "Page load timeout for " << sc.m_sFunctionName << "\n";
				SetComErrorMessage(sc.m_dwTimeoutMsgID, sc.m_dwHelpContextID);
				return HRES_LOAD_TIMEOUT_ERR;
			}
			else
			{
				ATLASSERT(VARIANT_FALSE == vbLoadTimeoutIsErr);
				return HRES_LOAD_TIMEOUT_WARN;
			}
		}
	}
	catch (const ExceptionServices::OperationNotAllowedException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_INVALID_OPERATION, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_OPERATION_NOT_APPLICABLE);
		return HRES_OPERATION_NOT_APPLICABLE;
	}
	catch (const ExceptionServices::InvalidParamException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(sc.m_dwInvalidParamMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_INVALID_ARG);
		return HRES_INVALID_ARG;
	}
	catch (const ExceptionServices::BrowserDisconnectedException& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(IDS_ERR_BROWSER_DISCONNECTED, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_BRWS_CONNECTION_LOST);
		return HRES_BRWS_CONNECTION_LOST_ERR;
	}
	catch (const ExceptionServices::Exception& except)
	{
		traceLog << except << "\n";
		SetComErrorMessage(sc.m_dwFailureMsgID, sc.m_dwHelpContextID);
		SetLastErrorCode(ERR_FAIL);
		return HRES_FAIL;
	}

	return HRES_OK;
}


BOOL BaseLibObject::IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID)
{
	if (m_spPlugin == NULL)
	{
		traceLog << "m_spPlugin is NULL in BaseLibObject::IsValidState\n";
		SetComErrorMessage(nPatternID, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return FALSE;
	}

	if (m_spCore == NULL)
	{
		traceLog << "m_spCore is NULL in BaseLibObject::IsValidState\n";
		SetComErrorMessage(nPatternID, szMethodName, dwHelpID);
		SetLastErrorCode(ERR_FAIL);
		return FALSE;
	}

	return TRUE;
}


void BaseLibObject::CloseBrowserPopups()
{
	if (NULL == m_spCore)
	{
		traceLog << "m_spCore is NULL in BaseLibObject::CloseBrowserPopups\n";
		return;
	}

	CCore* pCore = static_cast<CCore*>(m_spCore.p); // Down cast !!!
	if (pCore->GetAutoClosePopups() == FALSE)
	{
		traceLog << "pCore->GetAutoClosePopups() is FALSE in BaseLibObject::CloseBrowserPopups\n";
		return;
	}

	LONG nPluginThreadID = 0;
	if (!GetPluginThreadID(nPluginThreadID))
	{
		traceLog << "GetPluginThreadID failed in BaseLibObject::CloseBrowserPopups\n";
		return;
	}

	HWND hDialog = NULL;
	if (!Common::GetPopupByText(nPluginThreadID, _T(""), hDialog, NULL))
	{
		traceLog << "Common::GetPopupByText failed in BaseLibObject::CloseBrowserPopups nPluginThreadID=" << nPluginThreadID << "\n";
		return;
	}

	if (NULL == hDialog)
	{
		traceLog << "hDialog is NULL in BaseLibObject::CloseBrowserPopups\n";
		return;
	}

	String sPopupCaption = Common::GetWindowText(hDialog);

	int nButtonIndex = 0;
	if (s_autoClosePopupsMap.Lookup(sPopupCaption.c_str(), nButtonIndex))
	{
		BOOL bRes = Common::PressButtonOnPopup(hDialog, nButtonIndex + 1);
		UNREFERENCED_PARAMETER(bRes);
	}
	else
	{
		// Empty string means anything.
		if (s_autoClosePopupsMap.Lookup(_T(""), nButtonIndex))
		{
			BOOL bRes = Common::PressButtonOnPopup(hDialog, nButtonIndex + 1);
			UNREFERENCED_PARAMETER(bRes);
		}
	}
}


BOOL BaseLibObject::InitCloseBrowserPopups()
{
	// Race condition if two or more core objects are created concurrently (very unlikely?)
	if (s_bPopupsMapInit)
	{
		return FALSE;
	}

	s_bPopupsMapInit = TRUE;
	s_autoClosePopupsMap.RemoveAll();

	// English.
	s_autoClosePopupsMap.SetAt(_T("Information Bar"),        0);
	s_autoClosePopupsMap.SetAt(_T("Security Alert"),         0);
	s_autoClosePopupsMap.SetAt(_T("Security Warning"),       1);
	s_autoClosePopupsMap.SetAt(_T("AutoComplete Passwords"), 1);
	s_autoClosePopupsMap.SetAt(_T("AutoComplete"),           1); // WinXP + IE6
	s_autoClosePopupsMap.SetAt(_T("Webpage Error"),          1);
	s_autoClosePopupsMap.SetAt(_T("Error"),                  1); // WinXP + IE7
	s_autoClosePopupsMap.SetAt(_T("Internet Explorer"),      0); // IE6 WinXP, IE7 Vista

	// French.
	s_autoClosePopupsMap.SetAt(_T("Alerte de sécurité"),                        0);
	s_autoClosePopupsMap.SetAt(_T("Saisie semi-automatique des mots de passe"), 1);
	s_autoClosePopupsMap.SetAt(_T("Erreur de page Web"),                        1);

	// Italian.
	s_autoClosePopupsMap.SetAt(_T("Avviso di sicurezza"),               0);
	s_autoClosePopupsMap.SetAt(_T("Completamento automatico password"), 1);
	s_autoClosePopupsMap.SetAt(_T("Errore pagina Web"),                 1);

	// Spanish.
	s_autoClosePopupsMap.SetAt(_T("Alerta de seguridad"),       0);
	s_autoClosePopupsMap.SetAt(_T("Autocompletar contraseñas"), 1);
	s_autoClosePopupsMap.SetAt(_T("Error de página web"),       1);

	// German.
	s_autoClosePopupsMap.SetAt(_T("Sicherheitshinweis"),                   0);
	s_autoClosePopupsMap.SetAt(_T("AutoVervollständigen von Kennwörtern"), 1);
	s_autoClosePopupsMap.SetAt(_T("Webseitenfehler"),                      1);


	String sUserAutoClosePopupsFile;
	try
	{
		sUserAutoClosePopupsFile = Registry::RegGetStringValue(HKEY_CURRENT_USER, Settings::REG_SETTINGS_KEY_NAME, Settings::REG_SETTINGS_AUTO_CLOSE_POPUPS_FILE);
	}
	catch (const RegistryException&)
	{
	}

	String sMachineAutoClosePopupsFile;
	try
	{
		sMachineAutoClosePopupsFile = Registry::RegGetStringValue(HKEY_LOCAL_MACHINE, Settings::REG_SETTINGS_KEY_NAME, Settings::REG_SETTINGS_AUTO_CLOSE_POPUPS_FILE);
	}
	catch (const RegistryException&)
	{
	}

	BOOL bRes = TRUE;
	if (!sUserAutoClosePopupsFile.empty())
	{
		bRes = LoadAutoCloseSettings(sUserAutoClosePopupsFile);
	}

	if (!sMachineAutoClosePopupsFile.empty())
	{
		bRes = LoadAutoCloseSettings(sMachineAutoClosePopupsFile);
	}

	/*traceLog << "Auto-close popups settings in BaseLibObject::LoadAutoCloseSettings\n";

	POSITION pos = s_autoClosePopupsMap.GetStartPosition();
	while (TRUE)
	{
		CString sCrntKey;
		int     nCrntVal = 0;

		s_autoClosePopupsMap.GetNextAssoc(pos, sCrntKey, nCrntVal);

		USES_CONVERSION;
		traceLog << T2A((LPCTSTR)sCrntKey) << " : " << nCrntVal <<"\n";

		if (NULL == pos)
		{
			break;
		}
	}*/

	return bRes;
}


BOOL BaseLibObject::LoadAutoCloseSettings(const String& sXmlFilePath)
{
	CComQIPtr<IXMLDOMDocument> spXmlDoc;
	HRESULT hRes = spXmlDoc.CoCreateInstance(L"Msxml2.DOMDocument.3.0", NULL, CLSCTX_INPROC_SERVER);

	if (FAILED(hRes) || (spXmlDoc == NULL))
	{
		traceLog << "Cannot create XMl document in BaseLibObject::LoadAutoCloseSettings\n";
		return FALSE;
	}

	VARIANT_BOOL vbLoaded = VARIANT_FALSE;
	hRes = spXmlDoc->load(CComVariant(sXmlFilePath.c_str()), &vbLoaded);

	if (FAILED(hRes) || (VARIANT_FALSE == vbLoaded))
	{
		traceLog << "Cannot load XMl document in BaseLibObject::LoadAutoCloseSettings File:'" << sXmlFilePath << "'\n";
		return FALSE;
	}

	CComQIPtr<IXMLDOMElement> spRootElem;
	hRes = spXmlDoc->get_documentElement(&spRootElem);

	if (FAILED(hRes) || (spRootElem == NULL))
	{
		traceLog << "Cannot get root element in BaseLibObject::LoadAutoCloseSettings\n";
		return FALSE;
	}

	BOOL        bReplace = FALSE;
	CComVariant vAdd;

	hRes = spRootElem->getAttribute(CComBSTR("add"), &vAdd);
	if (SUCCEEDED(hRes) && (VT_BSTR == vAdd.vt) && (vAdd.bstrVal != NULL))
	{
		bReplace = (wcscmp(vAdd.bstrVal, L"1") != 0);
	}

	if (bReplace)
	{
		s_autoClosePopupsMap.RemoveAll();
	}

	CComQIPtr<IXMLDOMNodeList> spPopupList;
	hRes = spRootElem->selectNodes(CComBSTR(".//popup"), &spPopupList);

	if (FAILED(hRes))
	{
		traceLog << "Cannot selectNodes in BaseLibObject::LoadAutoCloseSettings\n";
		return FALSE;
	}

	while (TRUE)
	{
		CComQIPtr<IXMLDOMNode> spCrntNode;
		hRes = spPopupList->nextNode(&spCrntNode);

		CComQIPtr<IXMLDOMElement> sCrntElement = spCrntNode;
		if (sCrntElement == NULL)
		{
			// The end of the list.
			break;
		}

		CComVariant vCrntTitle;
		hRes = sCrntElement->getAttribute(CComBSTR("title"), &vCrntTitle);

		if (SUCCEEDED(hRes) && (VT_BSTR == vCrntTitle.vt) && (vCrntTitle.bstrVal != NULL))
		{
			CComVariant vCrntBtn;
			hRes = sCrntElement->getAttribute(CComBSTR("btnindex"), &vCrntBtn);
			if (SUCCEEDED(hRes) && (VT_BSTR == vCrntBtn.vt) && (vCrntBtn.bstrVal != NULL))
			{
				int  nCrntBtn = -1;
				BOOL bRes     = Common::StringToInt(vCrntBtn.bstrVal, &nCrntBtn);

				if (bRes)
				{
					s_autoClosePopupsMap.SetAt(vCrntTitle.bstrVal, nCrntBtn);
				}
			}
		}
	}

	return TRUE;
}


BOOL BaseLibObject::GetPluginThreadID(LONG& nThreadID)
{
	if (0 == m_nPluginThreadID)
	{
		if (m_spPlugin == NULL)
		{
			return FALSE;
		}

		HRESULT hRes = m_spPlugin->GetBrowserThreadID(&m_nPluginThreadID);
		if (FAILED(hRes))
		{
			traceLog << "m_spPlugin->GetBrowserThreadID failed in BaseLibObject::GetPluginThreadID\n";
			return FALSE;
		}
	}

	nThreadID = m_nPluginThreadID;
	return TRUE;
}


BOOL BaseLibObject::Sleep(DWORD dwMiliseconds, BOOL bDispatchMsg)
{
	LONG nPluginThreadID = 0;
	if (!GetPluginThreadID(nPluginThreadID))
	{
		traceLog << "GetPluginThreadID failed in BaseLibObject::Sleep\n";
		::Sleep(dwMiliseconds);

		return TRUE;
	}

	if (::GetCurrentThreadId() == nPluginThreadID)
	{
		if (bDispatchMsg)
		{
			// this object runs in the same thread as m_spPlugin
			// Here's a good place to dispatch messages so hosted browser controls won't block.
			// (e.g. winforms applications trying to automate a hosted WebBrowser control).

			DWORD dwStartTime = ::GetTickCount();
			while (TRUE)
			{
				MSG  msg = { 0 };
				BOOL bDispatched = FALSE;

				if (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
				{
					::DispatchMessage(&msg);
					bDispatched = TRUE;
				}

				DWORD dwCurrentTime = ::GetTickCount();
				if ((dwCurrentTime - dwStartTime) > dwMiliseconds)
				{
					// Timeout has expired. Quit the loop.
					break;
				}

				if (!bDispatched)
				{
					// No message to process yet; sleep for a short while.
					::Sleep(10);
				}
			}
		}

		return FALSE;
	}
	else
	{
		::Sleep(dwMiliseconds);
		return TRUE;
	}
}
