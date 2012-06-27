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

#pragma once

namespace HtmlHelpers
{
	BOOL                       IsWindowReady              (CComQIPtr<IHTMLWindow2> spWindow);
	BOOL                       IsBrowserReady             (CComQIPtr<IWebBrowser2> spBrws);
	CComQIPtr<IHTMLDocument2>  HtmlWindowToHtmlDocument   (CComQIPtr<IHTMLWindow2> spWindow);
	CComQIPtr<IWebBrowser2>    HtmlWindowToHtmlWebBrowser (CComQIPtr<IHTMLWindow2> spWindow);
	CComQIPtr<IHTMLWindow2>    HtmlWebBrowserToHtmlWindow (CComQIPtr<IWebBrowser2> spBrowser);
	CComQIPtr<IAccessible>     HtmlElementToAccessible    (CComQIPtr<IHTMLElement> spHtmlElement);
	CComQIPtr<IHTMLElement>    AccessibleToHtmlElement    (CComQIPtr<IAccessible> spAccessible);
	CComQIPtr<IHTMLWindow2>    AccessibleToHtmlWindow     (CComQIPtr<IAccessible> spAccessible);
	CComQIPtr<IHTMLDocument2>  AccessibleToHtmlDocument   (CComQIPtr<IAccessible> spAccessible);
	CComQIPtr<IAccessible>     HtmlDocumentToAccessible   (CComQIPtr<IHTMLDocument2> spDocument);
	CComQIPtr<IHTMLImgElement> GetAreaImage               (CComQIPtr<IHTMLAreaElement> spArea);
	CComQIPtr<IHTMLElement>    GetHtmlElementFromScreenPos(CComQIPtr<IHTMLDocument2> spDocument, long lX, long lY);
	BOOL                       IsComboControl             (CComQIPtr<IHTMLSelectElement> spSelectElem);
	BOOL                       IsMultipleListControl      (CComQIPtr<IHTMLSelectElement> spSelectElem);
	CComQIPtr<IHTMLElement>    GetAncestorByTag           (IHTMLElement* pElement, CComBSTR bstrTagName);

	CComBSTR GetTextAttributeValue    (CComQIPtr<IHTMLElement> spHtmlElement);
	BOOL     GetElemClickScreenPoint  (CComQIPtr<IHTMLElement> spElement, POINT* pClickPoint, BOOL bGetInputFileButton, int nZoomLevel);
	BOOL     GetAreaOffset            (CComQIPtr<IHTMLAreaElement> spArea, POINT* pOffset);
	BOOL     IsInputFileElement       (CComQIPtr<IHTMLElement> spHtmlElement);
	BOOL     GetElementScreenLocation (IHTMLElement* pElement, RECT* pRect, int nZoomLevel);
	HRESULT  GetIEServerWndFromElement(IHTMLElement* pElement, HWND* pIeServerWnd);
	HWND     GetIEWndFromBrowser      (CComQIPtr<IWebBrowser2> spBrowser);
	HWND     GetIEServerFromScrPt     (LONG x, LONG y);

	// Rise HTML events.
	HRESULT ClickElementAt    (IHTMLElement* pElement, LONG x, LONG y, BOOL bRightClick);
	BOOL    FireEventOnElement(CComQIPtr<IHTMLElement> spTargetElement, CComBSTR bstrEvName, WCHAR wchCharToRise = L'\0', LONG x = -1, LONG y = -1, BOOL bRightClick = FALSE);

	CComQIPtr<IHTMLElement> GetFrameElement          (CComQIPtr<IHTMLWindow2> spFrame);
	CComQIPtr<IWebBrowser2> GetBrowserFromIEServerWnd(HWND hIeWnd);
}
