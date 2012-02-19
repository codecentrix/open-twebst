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
#include "Keyboard.h"


namespace Keyboard
{
	const DWORD KEYSTROKE_PAUSE = 10;
	const DWORD SHIFT_KEY       =  1;
	const DWORD CTRL_KEY        =  2;
	const DWORD ALT_KEY         =  4;

	const DWORD SEND_KEY_DOWN = 0x1;
	const DWORD SEND_KEY_UP   = 0x2;
	const DWORD SEND_CHAR     = 0x4;


	BOOL SendCharacter(TCHAR tch);
	BOOL SendCharacter(TCHAR tch, HWND hIeWnd, BOOL bIsNormalEdit = TRUE);
	BOOL SendVirtualKey(TCHAR tch, DWORD vk, HKL hKL, HWND hIeWnd, UINT nFlags = SEND_KEY_DOWN | SEND_KEY_UP);
	BOOL SetCapsLockKeyState(BOOL bState);
	BOOL SendHomeKey(HWND hIeWnd);
	BOOL SendDelKey(HWND hIeWnd);
}

BOOL Keyboard::SetCapsLockKeyState(BOOL bState)
{
	BOOL bCapsOn = ::GetKeyState(VK_CAPITAL) & 0x0001;
	if (bCapsOn != bState)
	{
		// Set the caps lock key state by pressing it.
		INPUT capsLockInput[2];
		::ZeroMemory(capsLockInput, sizeof(INPUT) * 2);
		capsLockInput[0].type   = INPUT_KEYBOARD;
		capsLockInput[0].ki.wVk = VK_CAPITAL;

		capsLockInput[1].type       = INPUT_KEYBOARD;
		capsLockInput[1].ki.dwFlags = KEYEVENTF_KEYUP;
		capsLockInput[1].ki.wVk     = VK_CAPITAL;

		UINT nResult = ::SendInput(2, capsLockInput, sizeof(INPUT));
		return (nResult == 2);
	}

	return TRUE;
}


BOOL Keyboard::InputText(LPCTSTR szText)
{
	ATLASSERT(szText != NULL);

	BOOL bCapsOn = ::GetKeyState(VK_CAPITAL) & 0x0001;
	if (bCapsOn)
	{
		if (!SetCapsLockKeyState(FALSE))
		{
			traceLog << "SetCapsLockKeyState failed in Keyboard::InputText\n";
			return FALSE;
		}
	}

	size_t nLength = _tcslen(szText);
	for (size_t i = 0; i < nLength; ++i)
	{
		if (!SendCharacter(szText[i]))
		{
			traceLog << "SendCharacter failed in Keyboard::InputText\n";
			if (bCapsOn)
			{
				if (!SetCapsLockKeyState(TRUE))
				{
					traceLog << "SetCapsLockKeyState failed in Keyboard::InputText\n";
					return FALSE;
				}
			}
			return FALSE;
		}

		::Sleep(KEYSTROKE_PAUSE);
	}

	if (bCapsOn)
	{
		if (!SetCapsLockKeyState(TRUE))
		{
			traceLog << "SetCapsLockKeyState failed in Keyboard::InputText\n";
			return FALSE;
		}
	}

	return TRUE;
}


BOOL Keyboard::SendCharacter(TCHAR tch)
{
	INPUT inputs[2];
	INPUT shiftInputs[3];

	KEYBDINPUT keyInputDown;
	KEYBDINPUT keyInputUp;
	KEYBDINPUT keyShift;
	KEYBDINPUT keyAlt;
	KEYBDINPUT keyCtrl;

	HKL hKL = ::GetKeyboardLayout(::GetCurrentThreadId());

	short shVK = ::VkKeyScanEx(tch, hKL);
	if ((-1 == LOBYTE(shVK)) && (-1 == HIBYTE(shVK)))
	{
		traceLog << "VkKeyScanEx failed in Keyboard::SendCharacter for tch=" << tch << "\n";
		return FALSE;
	}

	// Press Shift, Ctrl and Alt keys /////////////////////////
	int  nShiftKeys = 0;
	BYTE shiftState = HIBYTE(shVK);
	if ((shiftState & SHIFT_KEY) != 0)
	{
		keyShift.dwFlags	= KEYEVENTF_EXTENDEDKEY;
		keyShift.wVk		= VK_SHIFT;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyShift;
		nShiftKeys++;
	}

	if ((shiftState & CTRL_KEY) != 0)
	{
		keyCtrl.dwFlags	= KEYEVENTF_EXTENDEDKEY;
		keyCtrl.wVk		= VK_CONTROL;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyCtrl;
		nShiftKeys++;
	}

	if ((shiftState & ALT_KEY) != 0)
	{
		keyAlt.dwFlags	= KEYEVENTF_EXTENDEDKEY;
		keyAlt.wVk		= VK_MENU;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyAlt;
		nShiftKeys++;
	}

	UINT nRes = 0;
	if (nShiftKeys > 0)
	{
		nRes = ::SendInput(nShiftKeys, shiftInputs, sizeof(INPUT));
		if (nShiftKeys != nRes)
		{
			traceLog << "SendInput failed in Keyboard::SendCharacter for tch=" << tch << " while sending the shift keys\n";
			return FALSE;
		}
	}

	// Press and release the normal key.
	keyInputDown.dwFlags	= KEYEVENTF_EXTENDEDKEY;
	keyInputDown.wVk		= (WORD)LOBYTE(shVK);

	keyInputUp.dwFlags	= KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;
	keyInputUp.wVk		= (WORD)LOBYTE(shVK);

	inputs[0].type = INPUT_KEYBOARD;
	inputs[0].ki = keyInputDown;

	inputs[1].type = INPUT_KEYBOARD;
	inputs[1].ki = keyInputUp;

	nRes = ::SendInput(2, inputs, sizeof(INPUT));
	if (nRes != 2)
	{
		traceLog << "SendInput failed in Keyboard::SendCharacter for tch=" << tch << "\n";
		return FALSE;
	}

	// Release Shift, Ctrl and Alt keys /////////////////////////
	nShiftKeys = 0;
	if ((shiftState & SHIFT_KEY) != 0)
	{
		keyShift.dwFlags	= KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;
		keyShift.wVk		= VK_SHIFT;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyShift;
		nShiftKeys++;
	}

	if ((shiftState & CTRL_KEY) != 0)
	{
		keyCtrl.dwFlags	= KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;
		keyCtrl.wVk		= VK_CONTROL;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyCtrl;
		nShiftKeys++;
	}

	if ((shiftState & ALT_KEY) != 0)
	{
		keyAlt.dwFlags	= KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;
		keyAlt.wVk		= VK_MENU;

		shiftInputs[nShiftKeys].type = INPUT_KEYBOARD;
		shiftInputs[nShiftKeys].ki = keyAlt;
		nShiftKeys++;
	}

	if (nShiftKeys > 0)
	{
		nRes = ::SendInput(nShiftKeys, shiftInputs, sizeof(INPUT));
		if (nShiftKeys != nRes)
		{
			traceLog << "SendInput failed in Keyboard::SendCharacter for tch=" << tch << " while realeasing the shift keys\n";
			return FALSE;
		}
	}

	return TRUE;
}


BOOL Keyboard::InputText(LPCTSTR szText, HWND hIeWnd, BOOL bIsNormalEdit)
{
	ATLASSERT(szText != NULL);
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	size_t nLength = _tcslen(szText);
	for (size_t i = 0; i < nLength; ++i)
	{
		if (!SendCharacter(szText[i], hIeWnd, bIsNormalEdit))
		{
			traceLog << "SendCharacter failed in Keyboard::InputText\n";
			return FALSE;
		}

		::Sleep(KEYSTROKE_PAUSE);
	}

	return TRUE;
}


BOOL Keyboard::SendCharacter(TCHAR tch, HWND hIeWnd, BOOL bIsNormalEdit)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	HKL   hKL  = ::GetKeyboardLayout(::GetCurrentThreadId());
	SHORT shVK = ::VkKeyScanEx(tch, hKL);
	if ((-1 == LOBYTE(shVK)) && (-1 == HIBYTE(shVK)))
	{
		traceLog << "VkKeyScanEx failed in Keyboard::SendCharacter for tch=" << tch << "\n";
		return FALSE;
	}

	UINT nFlags     = bIsNormalEdit ? SEND_CHAR : SEND_KEY_DOWN | SEND_KEY_UP | SEND_CHAR;
	UINT virtualKey = LOBYTE(shVK);
	return SendVirtualKey(tch, virtualKey, hKL, hIeWnd, nFlags);
}


BOOL Keyboard::SendVirtualKey(TCHAR tch, DWORD virtualKey, HKL hKL, HWND hIeWnd, UINT nFlags)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	if (0 == tch)
	{
		nFlags &= ~SEND_CHAR;
	}
	else if (Common::GetIEVersion() <= 6)
	{
		// Send only WM_CHAR for IE6.
		nFlags = SEND_CHAR;
	}

	const DWORD REPEAT_COUNT     =   0x0001;
	const DWORD SCAN_CODE        = 0xFF0000;
	const DWORD PREVIOUS_STATE   = 1 << 30;
	const DWORD TRANSITION_STATE = 1 << 31;

	UINT nScanCode = ::MapVirtualKeyEx(virtualKey, 0, hKL);
	if (!nScanCode)
	{
		traceLog << "Can not get the scan code in Keyboard::SendVirtualKey\n";
		return FALSE;
	}

	nScanCode = (nScanCode << 16) & SCAN_CODE;

	DWORD dwKeyDownLparam = REPEAT_COUNT | nScanCode;
	DWORD dwKeyUpLparam   = REPEAT_COUNT | PREVIOUS_STATE | TRANSITION_STATE | nScanCode;

	if (SEND_KEY_DOWN & nFlags)
	{
		BOOL bRes = ::PostMessage(hIeWnd, WM_KEYDOWN, (WPARAM)virtualKey , (LPARAM)dwKeyDownLparam);
		if (!bRes)
		{
			DWORD dwLastErr = ::GetLastError();
			traceLog << "PostMessage(WM_KEYDOWN) failed in Keyboard::SendVirtualKey, GetLastError=" << dwLastErr << "\n";
			return FALSE;
		}
	}

	if (SEND_CHAR & nFlags)
	{
		// Post wm_char message.
		BOOL bRes = ::PostMessage(hIeWnd, WM_CHAR, (WPARAM)tch, (LPARAM)nScanCode);
		if (!bRes)
		{
			DWORD dwLastErr = ::GetLastError();
			traceLog << "PostMessage(WM_CHAR) failed in Keyboard::SendVirtualKey, GetLastError=" << dwLastErr << "\n";
			return FALSE;
		}
	}

	if (SEND_KEY_UP & nFlags)
	{
		BOOL bRes =::PostMessage(hIeWnd, WM_KEYUP, (WPARAM)virtualKey , (LPARAM)dwKeyUpLparam);
		if (!bRes)
		{
			DWORD dwLastErr = ::GetLastError();
			traceLog << "PostMessage(WM_KEYUP) failed in Keyboard::SendVirtualKey, GetLastError=" << dwLastErr << "\n";
			return FALSE;
		}
	}

	return TRUE;
}


BOOL Keyboard::SendDelKey(HWND hIeWnd)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	HKL hKL  = ::GetKeyboardLayout(::GetCurrentThreadId());
	return SendVirtualKey(0, VK_DELETE, hKL, hIeWnd);
}


BOOL Keyboard::SendHomeKey(HWND hIeWnd)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	HKL hKL  = ::GetKeyboardLayout(::GetCurrentThreadId());
	return SendVirtualKey(0, VK_HOME, hKL, hIeWnd);
}


BOOL Keyboard::ClearInputFileControl(HWND hIeWnd, long nCharCount)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));
	ATLASSERT(nCharCount >= 0);

	if (!SendHomeKey(hIeWnd))
	{
		return FALSE;
	}

	for (int i = 0; i < nCharCount; ++i)
	{
		if (!SendDelKey(hIeWnd))
		{
			return FALSE;
		}
	}

	return TRUE;
}
