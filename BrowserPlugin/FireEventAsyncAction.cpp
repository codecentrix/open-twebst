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
#include "FireEventAsyncAction.h"
#include "HtmlHelpers.h"



FireEventAsyncAction::FireEventAsyncAction(CComQIPtr<IHTMLElement> spTargetElement, BSTR bstrEventName, LONG wCharToRise)
{
	ATLASSERT(spTargetElement != NULL);
	ATLASSERT(bstrEventName   != NULL);

	m_spTargetElement = spTargetElement;
	m_bstrEventName   = bstrEventName;
	m_wCharToRise     = wCharToRise;
}


HRESULT FireEventAsyncAction::DoAction()
{
	if ((NULL == m_spTargetElement) || (m_bstrEventName.Length() == 0))
	{
		return E_INVALIDARG;
	}

	BOOL bRes = HtmlHelpers::FireEventOnElement(m_spTargetElement, m_bstrEventName, (WCHAR)m_wCharToRise);
	if (!bRes)
	{
		traceLog << "Failed to rise onkeydown event in FireEventAsyncAction::DoAction\n";
	}

	return bRes ? S_OK : E_FAIL;
}


FireEventAsyncAction::~FireEventAsyncAction()
{
}
