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

// Common types and functions declarations.
#pragma once

// Constants, types, classes, functions used all over the project.
namespace Common
{
	// Types.
	typedef std::basic_string<TCHAR>        String;
	typedef std::basic_ostringstream<TCHAR> Ostringstream;

	struct DescriptorToken
	{
		enum TOKEN_OPERATOR { TOKEN_MATCH, TOKEN_NOT_MATCH };

		DescriptorToken(const String& sAttributeName, const String& sAttributePattern, DescriptorToken::TOKEN_OPERATOR tokenOp) :
			m_sAttributeName(sAttributeName),
			m_sAttributePattern(sAttributePattern),
			m_operator(tokenOp)
		{
		}

		String         m_sAttributeName;
		String         m_sAttributePattern;
		TOKEN_OPERATOR m_operator;
	};

	typedef BOOL (*FIND_CHILD_WND_CONDITION_CALLBACK)(HWND hCrntWnd, void*);

	// Constants.
	const DWORD START_XBIT_INJECT_TIMEOUT = 4000;
	const DWORD START_PROCESS_TIMEOUT     = 16000; // in miliseconds.
	const DWORD INTERNAL_GLOBAL_TIMEOUT = 60000; // in miliseconds.
	const DWORD INTERNAL_GLOBAL_PAUSE   = 50;    // in miliseconds.
	const DWORD TIME_SCALE              = 1;     // 1 = miliseconds, 10 = 1/100 seconds, 1000 = seconds.

	const TCHAR BROWSER_TITLE_ATTR_NAME[]        = _T("TITLE");
	const TCHAR BROWSER_URL_ATTR_NAME[]          = _T("URL");
	const TCHAR BROWSER_PID_ATTR_NAME[]          = _T("PID");
	const TCHAR BROWSER_TID_ATTR_NAME[]          = _T("TID");
	const TCHAR BROWSER_APP_NAME[]               = _T("APP");
	const TCHAR TEXT_SEARCH_ATTRIBUTE[]          = _T("UINAME");
	const TCHAR INDEX_SEARCH_ATTRIBUTE[]         = _T("INDEX");
	const TCHAR SEARCH_FRAME_ELEMENT_ATTR_NAME[] = _T("TWBSTSEARCHFRAMEELEMTICKCOUNT");

	// Global search flags passed from class library to browser plugin.
	const ULONG SEARCH_CHILDREN_ONLY   = 0x02;
	const ULONG SEARCH_ALL_HIERARCHY   =    0;
	const ULONG SEARCH_COLLECTION      = 0x04;
	const ULONG SEARCH_ELEMENT         =    0;
	const ULONG SEARCH_FRAME           = 0x08;
	const ULONG ADD_SELECTION          = 0x10;
	const ULONG SEARCH_HTML_DIALOG     = 0x20;
	const ULONG SEARCH_MODAL_HTML_DLG  = 0x40;
	const ULONG SEARCH_PARENT_ELEM     = 0x80;
	const ULONG SEARCH_SELECTED_OPTION = 0x100;
	const ULONG PERFORM_ASYNC_ACTION   = 0x200;

	// Extra info used to generate hardware events.
	const DWORD HARDWARE_EVENT_EXTRA_INFO = 0x2C1016BA;

	// Functions.
	//String  GetPluginRotName(DWORD dwThreadID);
	BOOL    GetWndProcName              (HWND hWnd, String& sOutProcName);
	HWND    GetTopParentWnd             (HWND hWnd);
	String  GetWndClass                 (HWND hWnd);
	String  TrimString                  (const String& s);
	String  TrimRightString             (const String& s);
	String  TrimLeftString              (const String& s);
	BOOL    MatchWildcardPattern        (LPCTSTR szPattern, LPCTSTR szText);
	BOOL    GetDescriptorTokensList     (SAFEARRAY* psa, std::list<DescriptorToken>* pTokens, int* pIndex = NULL);
	BOOL    IsValidSafeArray            (SAFEARRAY* psa);
	BOOL    SafeArraySize               (SAFEARRAY* psa, int* pSize);
	BOOL    IsEmptyOrBlank              (const CComBSTR& bstrText);
	void    StripLastSlash              (String* pStr);
	BOOL    IsValidOptionVariant        (const VARIANT& vOption);
	String  LoadStringFromRes           (UINT uID, HINSTANCE hInst = ::GetModuleHandle(NULL));
	int     GetIEVersion                ();
	BOOL    GetIEVersion                (CComBSTR& bstrVersion);
	HWND    GetChildWindowByClassName   (HWND hParentWindow, const String& sChildClassName, FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback = NULL, void* pObj = NULL);
	HWND    GetTopLevelWindowByClassName(HWND hIEWindow, const String& sChildClassName, FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback = NULL, void* pObj = NULL);
	HWND    GetTopLevelWindowByClassName(LONG nThreadID, const String& sChildClassName, FIND_CHILD_WND_CONDITION_CALLBACK pfCondCallback = NULL, void* pObj = NULL);
	BOOL    IsWindowsVistaOrLater       ();
	BOOL    StringToInt                 (const String& sNumber, int* pResult);

	inline String ToUpper(const String& sInputStr)
	{
		String sResult = sInputStr;
		std::transform(sResult.begin(), sResult.end(), sResult.begin(), (int(*)(int))_totupper);

		return sResult;
	}

	inline BOOL MatchWildcardPattern(const String& sPattern, const String& szText)
	{
		return MatchWildcardPattern(sPattern.c_str(), szText.c_str());
	}

	String GetCurrentProcessExeName();
	String GetCurrentModuleDir     (HMODULE hModule);

	String GetWindowText     (HWND hWnd);
	BOOL   GetPopupByText    (LONG nThreadID, const String& sText, HWND& hOutWnd, String* pPopupText);
	BOOL   PressButtonOnPopup(HWND hWndPopup, const String& sButtonText);
	BOOL   PressButtonOnPopup(HWND hWndPopup, int nIndex);
}


// Macros
#ifdef _UNICODE
	#define WIDEN2(x) L ## x
	#define WIDEN(x) WIDEN2(x)
	#define __TFILE__ WIDEN(__FILE__)	// __FILE__ is ASCII.
#else
	#define __TFILE__ __FILE__
#endif // _UNICODE
