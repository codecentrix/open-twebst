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
#include "SetFocusAwayAsyncAction.h"
#include "HtmlHelpers.h"



SetFocusAwayAsyncAction::SetFocusAwayAsyncAction(CComQIPtr<IHTMLElement> spTargetElement, BOOL bGenerateOnChange)
{
	ATLASSERT(spTargetElement != NULL);
	m_spTargetElement   = spTargetElement;
	m_bGenerateOnChange = bGenerateOnChange;
}


HRESULT SetFocusAwayAsyncAction::DoAction()
{
	if (NULL == m_spTargetElement)
	{
		return E_INVALIDARG;
	}

	const TCHAR* focusEvents[] = 
	{
		_T("onbeforedeactivate"), _T("onchange"), _T("ondeactivate"), _T("onfocusout"), _T("onblur"),
	};

	const int numberOfEvents = sizeof(focusEvents) / sizeof(focusEvents[0]);

	BOOL bRes = TRUE;
	for (int i = 0; i < numberOfEvents; ++i)
	{
		if ((!m_bGenerateOnChange) && !_tcsicmp(_T("onchange"), focusEvents[i]))
		{
			// Don't generate onchange event.
			continue;
		}

		BOOL bEventRised = HtmlHelpers::FireEventOnElement(m_spTargetElement, focusEvents[i]);
		if (!bEventRised)
		{
			bRes = FALSE;
			traceLog << "Failed to rise " << focusEvents[i] << " event in SetFocusAwayAsyncAction::DoAction\n";
		}
	}

	return bRes ? S_OK : E_FAIL;
}


SetFocusAwayAsyncAction::~SetFocusAwayAsyncAction()
{
}
