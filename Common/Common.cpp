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
#include "DebugServices.h"
#include "Common.h"
using namespace Common;


namespace Common
{
	struct FindChildInfo
	{
		FindChildInfo(const String& sClassNameToFind, FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback = NULL, void* pObj = NULL) :
			m_pfCondCallback(pfCondCallback),
			m_sChildClassName(sClassNameToFind),
			m_hFoundWnd(NULL),
			m_treahdIdToSearch(0),
			m_pCallbackParamObj(pObj)
		{
		}

		void*                             m_pCallbackParamObj;
		FIND_CHILD_WND_CONDITION_CALLBACK m_pfCondCallback;
		String                            m_sChildClassName;
		HWND                              m_hFoundWnd;
		DWORD                             m_treahdIdToSearch;
	};

	BOOL CALLBACK FindChildWndCallback(HWND hWnd, LPARAM lParam)
	{
		ATLASSERT(::IsWindow(hWnd));
		ATLASSERT(lParam != NULL);

		FindChildInfo* pFindInfo = (FindChildInfo*)lParam;

		if (pFindInfo->m_treahdIdToSearch != 0)
		{
			DWORD dwCrntThreadId = ::GetWindowThreadProcessId(hWnd, NULL);
			if (dwCrntThreadId != pFindInfo->m_treahdIdToSearch)
			{
				// Continue searching.
				return TRUE;
			}
		}

		if (Common::GetWndClass(hWnd) == pFindInfo->m_sChildClassName)
		{
			traceLog << "FindChildWndCallback: class name match " << pFindInfo->m_sChildClassName << "\n";

			if (pFindInfo->m_pfCondCallback != NULL)
			{
				if (pFindInfo->m_pfCondCallback(hWnd, pFindInfo->m_pCallbackParamObj))
				{
					traceLog << "FindChildWndCallback: pFindInfo->m_pfCondCallback is TRUE; return FALSE pFindInfo->m_hFoundWnd=" << hWnd << "\n";
					pFindInfo->m_hFoundWnd = hWnd;
					return FALSE;
				}
			}
			else
			{
				traceLog << "FindChildWndCallback: return FALSE pFindInfo->m_hFoundWnd=" << hWnd << "\n";
				pFindInfo->m_hFoundWnd = hWnd;
				return FALSE;
			}
		}

		return TRUE;
	}

	BOOL GetIEVersion(LPCTSTR szValName, CComBSTR& bstrVersion);
	int GetIEVersion(LPCTSTR szValName);
}


String Common::GetWndClass(HWND hWnd)
{
	if (!::IsWindow(hWnd))
	{
		return _T("");
	}

	const int MAX_CLASS_NAME_LEN = 128;
	TCHAR szClass[MAX_CLASS_NAME_LEN] = { 0 };
	const int nClassLen = sizeof(szClass) / sizeof(szClass[0]);

	int nResult = ::GetClassName(hWnd, szClass, nClassLen);
	_ASSERTE(nResult > 0);

	return szClass;
}


BOOL Common::GetDescriptorTokensList(SAFEARRAY* psa, std::list<DescriptorToken>* pTokens, int* pIndex)
{
	if ((!psa) || (::SafeArrayGetDim(psa) != 1))
	{
		return FALSE;
	}

	if (pIndex != NULL)
	{
		*pIndex = 0;
	}

	long    nLowerBound;
	HRESULT hRes = ::SafeArrayGetLBound(psa, 1, &nLowerBound);
	if (hRes != S_OK)
	{
		traceLog << "SafeArrayGetLBound failed in CExplorerPlugin::GetDescriptorTokensArray\n";
		return FALSE;
	}

	long nUpperBound;
	hRes = ::SafeArrayGetUBound(psa, 1, &nUpperBound);
	if (hRes != S_OK)
	{
		traceLog << "SafeArrayGetUBound failed in CExplorerPlugin::GetDescriptorTokensArray\n";
		return FALSE;
	}

	std::list<DescriptorToken> resultList;
	long nTokensNumber = nUpperBound - nLowerBound + 1;
	int  nIndex = 0;

	for (int i = 0; i < nTokensNumber; ++i)
	{
		long idx[] = { nLowerBound + i };
		VARIANT var;
		
		::VariantInit(&var);
		hRes = ::SafeArrayGetElement(psa, idx, (void*)&var);
		if (hRes != S_OK)
		{
			traceLog << "::SafeArrayGetElement(" << i << ") failed in CExplorerPlugin::GetDescriptorTokensArray\n";
			::VariantClear(&var);
			return FALSE;
		}

		if (var.vt != VT_BSTR)
		{
			traceLog << "Not a BSTR in the safearray at position " << i << " in CExplorerPlugin::GetDescriptorTokensArray\n";
			::VariantClear(&var);
			return FALSE;
		}

		USES_CONVERSION;
		String sToken = W2T(var.bstrVal);
		::VariantClear(&var);

		size_t nFirstEqual = sToken.find_first_of(_T('='));
		if (String::npos == nFirstEqual)
		{
			traceLog << "No = in token " << i << " in function CExplorerPlugin::GetDescriptorTokensList\n";
			return FALSE;
		}

		DescriptorToken::TOKEN_OPERATOR tokenOp;
		String sAttributeName = sToken.substr(0, nFirstEqual);

		if (sAttributeName.length() == 0)
		{
			traceLog << "Empty attribute name " << i << " in function CExplorerPlugin::GetDescriptorTokensList\n";
			return FALSE;
		}

		size_t nLastChar = sAttributeName.size() - 1;
		if (_T('!') == sAttributeName[nLastChar])
		{
			tokenOp = DescriptorToken::TOKEN_NOT_MATCH;
			sAttributeName.erase(nLastChar);
		}
		else
		{
			tokenOp = DescriptorToken::TOKEN_MATCH;
		}

		sAttributeName = TrimString(sAttributeName);

		if (!_tcsicmp(sAttributeName.c_str(), INDEX_SEARCH_ATTRIBUTE))
		{
			if (pIndex != NULL)
			{
				int    nIndex      = 0;
				String sIndex      = sToken.substr(nFirstEqual + 1);
				BOOL   bConvResult = StringToInt(sIndex, &nIndex);

				if (!bConvResult)
				{
					traceLog << "Invalid index '" << sIndex << "' in function CExplorerPlugin::GetDescriptorTokensList\n";
					return FALSE;
				}

				*pIndex = nIndex;
			}
		}
		else
		{
			String sPattern = Common::TrimString(sToken.substr(nFirstEqual + 1));
			resultList.push_back(DescriptorToken(sAttributeName, sPattern, tokenOp));
		}
	}

	pTokens->swap(resultList);
	return TRUE;
}


String Common::TrimLeftString(const String& s)
{
	size_t nFirst = s.find_first_not_of(_T(" \t\n\r"));
	if (String::npos != nFirst)
	{
		return s.substr(nFirst);
	}
	else
	{
		return _T("");
	}
}


String Common::TrimRightString(const String& s)
{
	size_t nLast = s.find_last_not_of(_T(" \t\n\r"));
	if (String::npos != nLast)
	{
		return s.substr(0, nLast + 1);
	}
	else
	{
		return _T("");
	}
}


String Common::TrimString(const String& s)
{
	return TrimLeftString(TrimRightString(s));
}


BOOL Common::IsValidSafeArray(SAFEARRAY* psa)
{
	if (NULL == psa)
	{
		traceLog << "psa is NULL in Common::IsValidSafeArray\n";
		return FALSE;
	}

	int nSafeArraySize;
	if (!SafeArraySize(psa, &nSafeArraySize))
	{
		traceLog << "Can not get the size of the safe array in Common::IsValidSafeArray\n";
		return FALSE;
	}

	traceLog << "Array size is " << nSafeArraySize << " in Common::IsValidSafeArray\n";
	return nSafeArraySize != 0;
}


BOOL Common::SafeArraySize(SAFEARRAY* psa, int* pSize)
{
	ATLASSERT(pSize != NULL);
	if (NULL == psa)
	{
		traceLog << "psa is NULL in Common::SafeArraySize\n";
		return FALSE;
	}

	if (::SafeArrayGetDim(psa) != 1)
	{
		traceLog << "psa has more than ONE dimension in Common::SafeArraySize\n";
		return FALSE;
	}

	long nLowerBound;
	HRESULT hRes = ::SafeArrayGetLBound(psa, 1, &nLowerBound);
	if (hRes != S_OK)
	{
		traceLog << "SafeArrayGetLBound failed in Common::SafeArraySize\n";
		return FALSE;
	}

	long nUpperBound;
	hRes = ::SafeArrayGetUBound(psa, 1, &nUpperBound);
	if (hRes != S_OK)
	{
		traceLog << "SafeArrayGetUBound failed in Common::SafeArraySize\n";
		return FALSE;
	}

	*pSize = nUpperBound - nLowerBound + 1;
	return TRUE;
}


BOOL Common::IsEmptyOrBlank(const CComBSTR& bstrText)
{
	ATLASSERT(bstrText != NULL);
	for (unsigned int i = 0; i < bstrText.Length(); ++i)
	{
		if (!::iswspace(bstrText[i]))
		{
			return FALSE;
		}
	}

	return TRUE;
}


void Common::StripLastSlash(String* pStr)
{
	ATLASSERT(pStr != NULL);
	size_t nStrLen = pStr->size();
	size_t nLastCharIndex = nStrLen - 1;

	if (0 == nStrLen)
	{
		// Empty string.
		return;
	}

	TCHAR ch = pStr->at(nLastCharIndex);
	if (pStr->at(nLastCharIndex) == _T('/'))
	{
		pStr->erase(nLastCharIndex);
	}
}


BOOL Common::IsValidOptionVariant(const VARIANT& vOption)
{
	BOOL bIsValidBstr   = ((VT_BSTR == vOption.vt) && (NULL != vOption.bstrVal));
	BOOL bIsValidNumber = (VT_I4  == vOption.vt) || (VT_I2  == vOption.vt) ||
	                      (VT_UI4 == vOption.vt) || (VT_UI2 == vOption.vt) ||
						  (VT_INT == vOption.vt);

	return bIsValidBstr || bIsValidNumber;
}


String Common::LoadStringFromRes(UINT uID, HINSTANCE hInst)
{
	int    nBufferSize = 128;
	TCHAR* pBuffer     = NULL;
	int    nPrevRes    = 0;

	while (true)
	{
		// Allocate memory for the buffer.
		pBuffer    = new TCHAR[nBufferSize];
		pBuffer[0] = _T('\0');

		// Load the string from resources.
		int nRes = ::LoadString(hInst, uID, pBuffer, nBufferSize);
		ATLASSERT(nRes !=0);

		if (0 == nRes)
		{
			// The uID string does not exist in the resources.
			break;
		}

		if (nRes < nBufferSize - 1)
		{
			// The buffer was large enough.
			break;
		}

		if (nPrevRes == nRes)
		{
			// There are no more characters to read from resources.
			break;
		}

		// Dealocate the buffer and increase the buffer size.
		delete[] pBuffer;
		pBuffer      = NULL;
		nPrevRes     = nRes;
		nBufferSize *= 2;
	}

	String sResult = pBuffer;
	delete[] pBuffer;
	pBuffer = NULL;

	return sResult;
}


BOOL Common::GetIEVersion(CComBSTR& bstrVersion)
{
	if (!GetIEVersion(_T("svcUpdateVersion"), bstrVersion))
	{
		if (!GetIEVersion(_T("svcVersion"), bstrVersion))
		{
			if (!GetIEVersion(_T("Version"), bstrVersion))
			{
				return FALSE;
			}
		}
	}

	return TRUE;
}


BOOL Common::GetIEVersion(LPCTSTR szValName, CComBSTR& bstrVersion)
{
	ATLASSERT(szValName != NULL);
	bstrVersion = _T("");

	HKEY hIEKey = NULL;
	LONG lRes   = ::RegOpenKeyEx(HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Internet Explorer"), 0, KEY_READ, &hIEKey);
	if (lRes != ERROR_SUCCESS)
	{
		// Can't open IE key. Assume is IE7.
		traceLog << "Can't open IE key in Common::GetIEVersion\n";
		return FALSE;
	}

	// Get the data from "Version" string value.
	const int BUFFER_VAL_SIZE = 128;
    TCHAR szVersion[BUFFER_VAL_SIZE] = { 0 };
	DWORD nValType = 0;
	DWORD nValLen  = BUFFER_VAL_SIZE * sizeof(TCHAR);        
	lRes = ::RegQueryValueEx(hIEKey, szValName, NULL, &nValType, (LPBYTE)szVersion, &nValLen);    

	if (lRes != ERROR_SUCCESS)
	{
		traceLog << "Can't get Version value in Common::GetIEVersion\n";
		::RegCloseKey(hIEKey);

		return FALSE;
	}

	::RegCloseKey(hIEKey);
	bstrVersion = szVersion;
	return TRUE;
}


int Common::GetIEVersion()
{
	static int nVersion = 0;
	if (nVersion != 0)
	{
		return nVersion;
	}

	nVersion = GetIEVersion(_T("svcUpdateVersion"));
	if (!nVersion)
	{
		nVersion = GetIEVersion(_T("svcVersion"));
		if (!nVersion)
		{
			nVersion = GetIEVersion(_T("Version"));
			if (!nVersion)
			{
				return 0;
			}
		}
	}

	return nVersion;
}


// Get version from registry. It is stored in 7.0.5730.11 format.
int Common::GetIEVersion(LPCTSTR szValName)
{
	ATLASSERT(szValName != NULL);

	HKEY hIEKey = NULL;
	LONG lRes   = ::RegOpenKeyEx(HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Internet Explorer"), 0, KEY_READ, &hIEKey);
	if (lRes != ERROR_SUCCESS)
	{
		traceLog << "Can't open IE key in Common::GetIEVersion\n";
		return 0;
	}

	// Get the data from "Version" string value.
	const int BUFFER_VAL_SIZE = 128;
    TCHAR szVersion[BUFFER_VAL_SIZE] = { 0 };
	DWORD nValType = 0;
	DWORD nValLen  = BUFFER_VAL_SIZE * sizeof(TCHAR);        
	lRes = ::RegQueryValueEx(hIEKey, szValName, NULL, &nValType, (LPBYTE)szVersion, &nValLen);    

	if (lRes != ERROR_SUCCESS)
	{
		traceLog << "Can't get Version value in Common::GetIEVersion\n";
		::RegCloseKey(hIEKey);

		return 0;
	}

	::RegCloseKey(hIEKey);

	// Find first "."
	TCHAR* pPoint = _tcschr(szVersion, _T('.'));
	if (pPoint != NULL)
	{
		*pPoint = _T('\0');
	}
	else
	{
		traceLog << "Malformed Version registry value in Common::GetIEVersion\n";
	}

    int nVersion = _ttoi(szVersion);
	if (0 == nVersion)
	{
		// Invalid version. Assume IE7.
		traceLog << "Invalid Version string in Common::GetIEVersion\n";
	}

	return nVersion;
}


HWND Common::GetChildWindowByClassName(HWND hParentWindow, const String& sChildClassName, FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback, void* pObj)
{
	ATLASSERT(sChildClassName.length() > 0);

	if (!::IsWindow(hParentWindow))
	{
		return NULL;
	}

	FindChildInfo findInfo(sChildClassName, pfCondCallback, pObj);
	::EnumChildWindows(hParentWindow, FindChildWndCallback, (LPARAM)(&findInfo));

	return findInfo.m_hFoundWnd;
}


HWND Common::GetTopLevelWindowByClassName
	(
		HWND                              hIEWindow,
		const String&                     sChildClassName,
		FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback,
		void*                             pObj
	)
{
	ATLASSERT(sChildClassName.length() > 0);

	if (!::IsWindow(hIEWindow))
	{
		return NULL;
	}

	DWORD dwThreadID = ::GetWindowThreadProcessId(hIEWindow, NULL);
	return GetTopLevelWindowByClassName(dwThreadID, sChildClassName, pfCondCallback, pObj);
}


HWND Common::GetTopLevelWindowByClassName
	(
		LONG                              nThreadID,
		const String&                     sChildClassName,
		FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback,
		void*                             pObj
	)
{
	ATLASSERT(sChildClassName.length() > 0);
	ATLASSERT(nThreadID != 0);

	traceLog << "GetTopLevelWindowByClassName begins with nThreadID=" << nThreadID << " and sChildClassName=" << sChildClassName << "\n";

	FindChildInfo findInfo(sChildClassName, pfCondCallback, pObj);
	findInfo.m_treahdIdToSearch = nThreadID;

	::EnumWindows(FindChildWndCallback, (LPARAM)(&findInfo));

	traceLog << "GetTopLevelWindowByClassName ends with findInfo.m_hFoundWnd=" << findInfo.m_hFoundWnd << "\n";
	return findInfo.m_hFoundWnd;
}


BOOL Common::IsWindowsVistaOrLater()
{
	OSVERSIONINFO osVer;
	osVer.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

	BOOL bErr = ::GetVersionEx(&osVer);
	if (!bErr)
	{
		return FALSE;
	}

	return (osVer.dwMajorVersion >= 6) ;
}


HWND Common::GetTopParentWnd(HWND hWnd)
{
	if (!::IsWindow(hWnd))
	{
		return NULL;
	}

	// Find top-level parent window.
	HWND hTopWnd = hWnd;

	while (TRUE)
	{
		HWND hParentWnd = ::GetParent(hTopWnd);
		if (NULL == hParentWnd)
		{
			break;
		}

		hTopWnd = hParentWnd;
	}

	return hTopWnd;
}


String Common::GetCurrentProcessExeName()
{
	// Get the name of the executable that runs the code.
	TCHAR szCurrentExeName[MAX_PATH + 1] = { 0 };

	::GetModuleFileName(NULL, szCurrentExeName, MAX_PATH);
	szCurrentExeName[MAX_PATH] = _T('\0');

	// Get the short exe name.
	TCHAR* szShortName = _tcsrchr(szCurrentExeName, _T('\\'));
	szShortName = (szShortName ? szShortName + 1 : szCurrentExeName);

	return szShortName;
}


String Common::GetCurrentModuleDir(HMODULE hModule)
{
	// Get the name of the hModule.
	TCHAR szCurrentExeName[MAX_PATH + 1] = { 0 };

	::GetModuleFileName(hModule, szCurrentExeName, MAX_PATH);
	szCurrentExeName[MAX_PATH] = _T('\0');

	// Get the short exe name.
	TCHAR* szDirNameEnd = _tcsrchr(szCurrentExeName, _T('\\'));
	if (szDirNameEnd != NULL)
	{
		*szDirNameEnd = _T('\0');
	}

	return szCurrentExeName;
}


BOOL Common::StringToInt(const String& sNumber, int* pResult)
{
	ATLASSERT(pResult != NULL);

	for (size_t i = 0; i < sNumber.size(); ++i)
	{
		if (!_istdigit(sNumber[i]))
		{
			return FALSE;
		}
	}

	int nResult = _stscanf_s(sNumber.c_str(), _T("%d"), pResult);
	return (1 == nResult);
}


//////////////////////////////// IE Popup /////////////////////////////////////
namespace Common
{
	struct FindPopupCondition
	{
		FindPopupCondition(const String& sText)
		{
			m_sText = TrimString(sText);
		}

		String       m_sText;
		String       m_sPopupTextResult;
	};

	BOOL FindPopupCallBack        (HWND hCrntWnd, void* pObjCond);
	BOOL FindStaticByTextCallback (HWND hCrntWnd, void* pObjCond);
	BOOL FindButtonByTextCallBack (HWND hCrntWnd, void* pObjCond);
	BOOL FindButtonByIndexCallBack(HWND hCrntBtn, void* pObjCond);
}


BOOL Common::GetPopupByText(LONG nThreadID, const String& sText, HWND& hOutWnd, String* pPopupText)
{
	hOutWnd = NULL;

	FindPopupCondition popupCond(sText);
	hOutWnd = GetTopLevelWindowByClassName(nThreadID, _T("#32770"), FindPopupCallBack, &popupCond);
	if (pPopupText != NULL)
	{
		*pPopupText = popupCond.m_sPopupTextResult;
	}

	return TRUE;
}


BOOL Common::FindPopupCallBack(HWND hCrntWnd, void* pObjCond)
{
	HWND hStatic = GetChildWindowByClassName(hCrntWnd, _T("Static"), FindStaticByTextCallback, pObjCond);

	traceLog << "FindPopupCallBack hStatic=" << hStatic << "\n";
	return (hStatic != NULL);
}


BOOL Common::FindStaticByTextCallback(HWND hCrntWnd, void* pObjCond)
{
	ATLASSERT(pObjCond != NULL);
	traceLog << "FindStaticByTextCallback begins with hCrntWnd=" << hCrntWnd << "\n";

	FindPopupCondition* pCond = static_cast<FindPopupCondition*>(pObjCond);

	// Get the text of the current "Static" child control.
	String sText = TrimString(GetWindowText(hCrntWnd));

	if (pCond->m_sText == _T(""))
	{
		// Empty text means any popup.
		traceLog << "FindStaticByTextCallback returns TRUE as pCond->m_sText is empty\n";
		pCond->m_sPopupTextResult = sText;

		return TRUE;
	}
	else
	{
		BOOL bMatch = Common::MatchWildcardPattern(pCond->m_sText, sText);
		if (bMatch)
		{
			pCond->m_sPopupTextResult = sText;
		}

		return bMatch;
	}
}



String Common::GetWindowText(HWND hWnd)
{
	// Get the lenght of the window's text.
    int    nLength   = ::GetWindowTextLength(hWnd);
    LPTSTR szWndText = new TCHAR[nLength + 1];

    ::GetWindowText(hWnd, szWndText, nLength + 1);

    String sResult = szWndText;
    delete []szWndText;
    
    return sResult;
}


BOOL Common::FindButtonByTextCallBack(HWND hCrntBtn, void* pObjCond)
{
	ATLASSERT(pObjCond != NULL);

	// Get the style of the button.
	LONG btnstyle = GetWindowLong(hCrntBtn, GWL_STYLE) & BS_TYPEMASK;
	if ((btnstyle != BS_PUSHBUTTON) && (btnstyle != BS_DEFPUSHBUTTON))
	{
		return FALSE;
	}

	String* pTextToFind   = static_cast<String*>(pObjCond);
	if (pTextToFind->empty())
	{
		// Any button will do it.
		return TRUE;
	}
	else
	{
		String sButtonString = GetWindowText(hCrntBtn);
		return (sButtonString == *pTextToFind);
	}
}


BOOL Common::PressButtonOnPopup(HWND hWndPopup, const String& sButtonText)
{
	HWND hButton = GetChildWindowByClassName(hWndPopup, _T("Button"), FindButtonByTextCallBack,(void*)(&sButtonText));
	if (::IsWindow(hButton))
	{
		::SendMessage(hButton, BM_CLICK, 0, 0);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


BOOL Common::FindButtonByIndexCallBack(HWND hCrntBtn, void* pObjCond)
{
	// Get the style of the button.
	LONG btnstyle = GetWindowLong(hCrntBtn, GWL_STYLE) & BS_TYPEMASK;
	if ((btnstyle != BS_PUSHBUTTON) && (btnstyle != BS_DEFPUSHBUTTON))
	{
		// Count only push buttons and not radios, checkboxes etc.
		return FALSE;
	}

	ATLASSERT(pObjCond != NULL);
	int* pIndex = static_cast<int*>(pObjCond);

	*pIndex = *pIndex - 1;

	if (*pIndex > 0)
	{
		return FALSE;
	}
	else
	{
		return TRUE;
	}
}


BOOL Common::PressButtonOnPopup(HWND hWndPopup, int nIndex)
{
	if (nIndex <= 0)
	{
		return FALSE;
	}

	HWND hButton = GetChildWindowByClassName(hWndPopup, _T("Button"), FindButtonByIndexCallBack,(void*)(&nIndex));
	if (::IsWindow(hButton))
	{
		::SendMessage(hButton, BM_CLICK, 0, 0);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


BOOL Common::GetWndProcName(HWND hWnd, String& sOutProcName)
{
	if (!::IsWindow(hWnd))
	{
		return FALSE;
	}

	DWORD dwProcId = 0;
	::GetWindowThreadProcessId(hWnd, &dwProcId);

	if (!dwProcId)
	{
		traceLog << "GetWindowThreadProcessId failed in Common::GetWndProcName\n";
		return FALSE;
	}

	HANDLE hProc = ::OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, dwProcId);
	if (NULL == hProc)
	{
		traceLog << "OpenProcess failed in Common::GetWndProcName\n";
		return FALSE;
	}

	TCHAR szProcPath[MAX_PATH + 1] = { 0 };
	if (!::GetModuleFileNameEx(hProc, NULL, szProcPath, MAX_PATH))
	{
		traceLog << "GetModuleFileNameEx failed in Common::GetWndProcName\n";

		::CloseHandle(hProc);
		hProc = NULL;

		return FALSE;
	}

	TCHAR* szResult       = szProcPath;
	TCHAR* pLastBackSlash = _tcsrchr(szProcPath, _T('\\'));

	if (pLastBackSlash != NULL)
	{
		szResult = pLastBackSlash + 1;
	}

	::CloseHandle(hProc);
	hProc = NULL;

	sOutProcName = szResult;
	return TRUE;
}


// TODO: add support for wildcard escape ** for * and ?? for ?
BOOL Common::MatchWildcardPattern(LPCTSTR szPattern, LPCTSTR szText)
{
	if (!szPattern && !szText)
	{
		// NULL matching NULL.
		return TRUE;
	}
	else
	{
		if (!szPattern || !szText)
		{
			// NULL pattern cannot match non-NULL text.
			// non-NULL pattern cannot match NULL text.
			return FALSE;
		}
	}

	while (*szPattern && *szText)
	{
		switch (*szPattern)
		{
			case _T('*'):
			{
				if (*(szPattern + 1) == _T('\0'))
				{
					return TRUE;
				}

				if (MatchWildcardPattern(szPattern + 1, szText) == FALSE)
				{
					szText++;
				}
				else
				{
					return TRUE;
				}

				break;
			}

			case _T('?'):
			{
				szPattern++;
				szText++;
				break;
			}

			default:
			{
				if (*szPattern == *szText)
				{
					szPattern++;
					szText++;
				}
				else
				{
					return FALSE;
				}
			}
		}
	}

	return (*szText == 0) && (*szPattern == 0 || ((*szPattern == _T('*') && *(szPattern + 1) == 0)));
}
