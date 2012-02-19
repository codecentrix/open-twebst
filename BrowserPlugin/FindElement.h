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
#include "LocalElementCollection.h"
using namespace std;

namespace FindElementAlgorithms
{
	// Functions that search html elements.
	CComQIPtr<IHTMLElement> FindParentElement (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindElement       (CComQIPtr<IWebBrowser2> spBrowser, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindElement       (CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindChildElement  (CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindElement       (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindChildElement  (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<IHTMLElement> FindSelectedOption(CComQIPtr<IHTMLSelectElement> spSelectElement);

	// Functions that search collections of elements.
	CComQIPtr<ILocalElementCollection> FindAllSelectedOptions(CComQIPtr<IHTMLSelectElement> spSelectElement);
	CComQIPtr<ILocalElementCollection> FindAllElements       (CComQIPtr<IWebBrowser2> spBrowser, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<ILocalElementCollection> FindAllElements       (CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<ILocalElementCollection> FindChildrenElements  (CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<ILocalElementCollection> FindAllElements       (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);
	CComQIPtr<ILocalElementCollection> FindChildrenElements  (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa);

	list<DWORD> FindOptionsIndexes(CComQIPtr<IHTMLElement> spSelectElement, BSTR bstrCondition);
}
