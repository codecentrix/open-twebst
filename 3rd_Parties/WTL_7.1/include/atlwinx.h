// Windows Template Library - WTL version 7.1
// Copyright (C) 1997-2003 Microsoft Corporation
// All rights reserved.
//
// This file is a part of the Windows Template Library.
// The code and information is provided "as-is" without
// warranty of any kind, either expressed or implied.

#ifndef __ATLWINX_H__
#define __ATLWINX_H__

#pragma once

#ifndef __cplusplus
	#error ATL requires C++ compilation (use a .cpp suffix)
#endif

#ifndef __ATLAPP_H__
	#error atlwinx.h requires atlapp.h to be included first
#endif

#if (_ATL_VER >= 0x0700)
  #include <atlwin.h>
#endif //(_ATL_VER >= 0x0700)


///////////////////////////////////////////////////////////////////////////////
// Classes in this file:
//
// _U_RECT
// _U_MENUorID
// _U_STRINGorID


///////////////////////////////////////////////////////////////////////////////
// Command Chaining Macros

#define CHAIN_COMMANDS(theChainClass) \
	if(uMsg == WM_COMMAND) \
		CHAIN_MSG_MAP(theChainClass)

#define CHAIN_COMMANDS_ALT(theChainClass, msgMapID) \
	if(uMsg == WM_COMMAND) \
		CHAIN_MSG_MAP_ALT(theChainClass, msgMapID)

#define CHAIN_COMMANDS_MEMBER(theChainMember) \
	if(uMsg == WM_COMMAND) \
		CHAIN_MSG_MAP_MEMBER(theChainMember)

#define CHAIN_COMMANDS_ALT_MEMBER(theChainMember, msgMapID) \
	if(uMsg == WM_COMMAND) \
		CHAIN_MSG_MAP_ALT_MEMBER(theChainMember, msgMapID)


///////////////////////////////////////////////////////////////////////////////
// Reflected message handler macros for message maps (for ATL 3.0)

#if (_ATL_VER < 0x0700)

#define REFLECTED_COMMAND_HANDLER(id, code, func) \
	if(uMsg == OCM_COMMAND && id == LOWORD(wParam) && code == HIWORD(wParam)) \
	{ \
		bHandled = TRUE; \
		lResult = func(HIWORD(wParam), LOWORD(wParam), (HWND)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_COMMAND_ID_HANDLER(id, func) \
	if(uMsg == OCM_COMMAND && id == LOWORD(wParam)) \
	{ \
		bHandled = TRUE; \
		lResult = func(HIWORD(wParam), LOWORD(wParam), (HWND)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_COMMAND_CODE_HANDLER(code, func) \
	if(uMsg == OCM_COMMAND && code == HIWORD(wParam)) \
	{ \
		bHandled = TRUE; \
		lResult = func(HIWORD(wParam), LOWORD(wParam), (HWND)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_COMMAND_RANGE_HANDLER(idFirst, idLast, func) \
	if(uMsg == OCM_COMMAND && LOWORD(wParam) >= idFirst  && LOWORD(wParam) <= idLast) \
	{ \
		bHandled = TRUE; \
		lResult = func(HIWORD(wParam), LOWORD(wParam), (HWND)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_COMMAND_RANGE_CODE_HANDLER(idFirst, idLast, code, func) \
	if(uMsg == OCM_COMMAND && code == HIWORD(wParam) && LOWORD(wParam) >= idFirst  && LOWORD(wParam) <= idLast) \
	{ \
		bHandled = TRUE; \
		lResult = func(HIWORD(wParam), LOWORD(wParam), (HWND)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_NOTIFY_HANDLER(id, cd, func) \
	if(uMsg == OCM_NOTIFY && id == ((LPNMHDR)lParam)->idFrom && cd == ((LPNMHDR)lParam)->code) \
	{ \
		bHandled = TRUE; \
		lResult = func((int)wParam, (LPNMHDR)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_NOTIFY_ID_HANDLER(id, func) \
	if(uMsg == OCM_NOTIFY && id == ((LPNMHDR)lParam)->idFrom) \
	{ \
		bHandled = TRUE; \
		lResult = func((int)wParam, (LPNMHDR)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_NOTIFY_CODE_HANDLER(cd, func) \
	if(uMsg == OCM_NOTIFY && cd == ((LPNMHDR)lParam)->code) \
	{ \
		bHandled = TRUE; \
		lResult = func((int)wParam, (LPNMHDR)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_NOTIFY_RANGE_HANDLER(idFirst, idLast, func) \
	if(uMsg == OCM_NOTIFY && ((LPNMHDR)lParam)->idFrom >= idFirst && ((LPNMHDR)lParam)->idFrom <= idLast) \
	{ \
		bHandled = TRUE; \
		lResult = func((int)wParam, (LPNMHDR)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#define REFLECTED_NOTIFY_RANGE_CODE_HANDLER(idFirst, idLast, cd, func) \
	if(uMsg == OCM_NOTIFY && cd == ((LPNMHDR)lParam)->code && ((LPNMHDR)lParam)->idFrom >= idFirst && ((LPNMHDR)lParam)->idFrom <= idLast) \
	{ \
		bHandled = TRUE; \
		lResult = func((int)wParam, (LPNMHDR)lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

#endif //(_ATL_VER < 0x0700)


///////////////////////////////////////////////////////////////////////////////
// Dual argument helper classes (for ATL 3.0)

#if (_ATL_VER < 0x0700)

namespace ATL
{

class _U_RECT
{
public:
	_U_RECT(LPRECT lpRect) : m_lpRect(lpRect)
	{ }
	_U_RECT(RECT& rc) : m_lpRect(&rc)
	{ }
	LPRECT m_lpRect;
};

class _U_MENUorID
{
public:
	_U_MENUorID(HMENU hMenu) : m_hMenu(hMenu)
	{ }
	_U_MENUorID(UINT nID) : m_hMenu((HMENU)LongToHandle(nID))
	{ }
	HMENU m_hMenu;
};

class _U_STRINGorID
{
public:
	_U_STRINGorID(LPCTSTR lpString) : m_lpstr(lpString)
	{ }
	_U_STRINGorID(UINT nID) : m_lpstr(MAKEINTRESOURCE(nID))
	{ }
	LPCTSTR m_lpstr;
};

}; //namespace ATL

#endif //(_ATL_VER < 0x0700)


namespace WTL
{

///////////////////////////////////////////////////////////////////////////////
// Forward notifications support for message maps (for ATL 3.0)

#if (_ATL_VER < 0x0700)

// forward notifications support
#define FORWARD_NOTIFICATIONS() \
	{ \
		bHandled = TRUE; \
		lResult = WTL::Atl3ForwardNotifications(m_hWnd, uMsg, wParam, lParam, bHandled); \
		if(bHandled) \
			return TRUE; \
	}

static LRESULT Atl3ForwardNotifications(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	LRESULT lResult = 0;
	switch(uMsg)
	{
	case WM_COMMAND:
	case WM_NOTIFY:
#ifndef _WIN32_WCE
	case WM_PARENTNOTIFY:
#endif //!_WIN32_WCE
	case WM_DRAWITEM:
	case WM_MEASUREITEM:
	case WM_COMPAREITEM:
	case WM_DELETEITEM:
	case WM_VKEYTOITEM:
	case WM_CHARTOITEM:
	case WM_HSCROLL:
	case WM_VSCROLL:
	case WM_CTLCOLORBTN:
	case WM_CTLCOLORDLG:
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORLISTBOX:
	case WM_CTLCOLORMSGBOX:
	case WM_CTLCOLORSCROLLBAR:
	case WM_CTLCOLORSTATIC:
		lResult = ::SendMessage(::GetParent(hWnd), uMsg, wParam, lParam);
		break;
	default:
		bHandled = FALSE;
		break;
	}
	return lResult;
}

#endif //(_ATL_VER < 0x0700)

}; //namespace WTL

#endif // __ATLWINX_H__
