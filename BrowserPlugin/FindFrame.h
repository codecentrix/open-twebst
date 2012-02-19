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
using namespace std;

namespace FindFrameAlgorithms
{
	// Functions that search html elements.
	CComQIPtr<IHTMLWindow2> FindFrame          (CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa);
	CComQIPtr<IHTMLWindow2> FindFrame          (CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa);
	CComQIPtr<IHTMLWindow2> FindChildFrame     (CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa);
	CComQIPtr<IHTMLWindow2> FindHtmlDialog     (CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa);
	CComQIPtr<IHTMLWindow2> FindModalHtmlDialog(CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa);


	// Functions that search collections of elements.
	list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindAllFrames     (CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa, LONG nSearchFlags);
	list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindAllFrames     (CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa, LONG nSearchFlags);
	list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindChildrenFrames(CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa, LONG nSearchFlags);
}
