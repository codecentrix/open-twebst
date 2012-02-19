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
#include "Exceptions.h"
#include "DebugServices.h"
#include "SearchFrameVisitor.h"
#include "HtmlHelpers.h"
using namespace std;
using namespace Common;


namespace FindFrameAlgorithms
{
	// Constructor.
	SearchFrameVisitor::SearchFrameVisitor(SAFEARRAY* psa)
	{
		if (!Common::IsValidSafeArray(psa))
		{
			throw CreateInvalidParamException(_T("Invalid safearray parameter in SearchFrameVisitor constructor"));
		}

		BOOL bRes = Common::GetDescriptorTokensList(psa, &m_tokens, &m_nSearchIndex);
		if (!bRes)
		{
			throw CreateInvalidParamException(_T("Can NOT get the list of descriptor tokens in SearchFrameVisitor constructor"));
		}
	}


	SearchFrameVisitor::~SearchFrameVisitor(void)
	{
	}


	BOOL SearchFrameVisitor::VisitFrame(CComQIPtr<IHTMLWindow2> spFrame)
	{
		ATLASSERT(spFrame != NULL);

		for (list<DescriptorToken>::iterator itToken = m_tokens.begin(); itToken != m_tokens.end(); ++itToken)
		{
			// Get the value of the visited element for the current attribute.
			String sAttributeValue;
			BOOL   bRes = GetFrameAttributeValue(spFrame, itToken->m_sAttributeName, &sAttributeValue);
			if (FALSE == bRes)
			{
				traceLog << "Can NOT get the value of the current frame " << itToken->m_sAttributeName << " in SearchFrameVisitor::VisitFrame no regexp case\n";
				return TRUE;	// Continue seraching.
			}

			// For SRC attribute ignore the last slash.
			String sAttributePattern = itToken->m_sAttributePattern;
			if (!_tcsicmp(itToken->m_sAttributeName.c_str(), _T("src")))
			{
				StripLastSlash(&sAttributePattern);
				StripLastSlash(&sAttributeValue);
			}

			BOOL bMatch = Common::MatchWildcardPattern(sAttributePattern, sAttributeValue);
			if (DescriptorToken::TOKEN_MATCH == itToken->m_operator)
			{
				if (!bMatch)
				{
					return TRUE; // Continue searching.
				}
			}
			else
			{
				ATLASSERT(DescriptorToken::TOKEN_NOT_MATCH == itToken->m_operator);
				if (bMatch)
				{
					return TRUE; // Continue searching.
				}
			}
		}

		ATLASSERT(m_nSearchIndex >= 0);

		if (0 == m_nSearchIndex)
		{
			// Stop searching. Attributes and index matching.
			m_spFoundFrame = spFrame;
			return FALSE;
		}
		else
		{
			// Index does NOT match. Continue searching.
			m_nSearchIndex--;
			return TRUE;
		}
	}


	CComQIPtr<IHTMLWindow2> SearchFrameVisitor::GetFoundHtmlFrame()
	{
		return m_spFoundFrame;
	}


	BOOL SearchFrameVisitor::GetFrameAttributeValue(CComQIPtr<IHTMLWindow2> spFrame,
	                                                  const String& sAttributeName,
                                                      String* pAttributeValue)
	{
		ATLASSERT(spFrame != NULL);
		ATLASSERT(pAttributeValue != NULL);

		CComQIPtr<IHTMLFrameBase> spFrameBase = HtmlHelpers::GetFrameElement(spFrame);
		if (spFrameBase == NULL)
		{
			// It is a top level window, it can NOT be converted to an element.
			if (_tcsicmp(sAttributeName.c_str(), _T("src")))
			{
				traceLog << "It should be the top window. For the top window only SRC attribute is valid\n";
				return FALSE;
			}
			else
			{
				CComQIPtr<IHTMLLocation> spLocation;
				HRESULT hRes = spFrame->get_location(&spLocation);
				if ((spLocation != NULL) && (S_OK == hRes))
				{
					CComBSTR bstrLocation;
					hRes = spLocation->get_href(&bstrLocation);
					if ((S_OK == hRes) && (bstrLocation != NULL))
					{
						USES_CONVERSION;
						*pAttributeValue = Common::TrimString(W2T(bstrLocation));
						return TRUE;
					}
				}

				traceLog << "Can NOT get the SRC attribute for the top window in SearchFrameVisitor::GetFrameAttributeValue\n";
				return FALSE;
			}
		}


		CComQIPtr<IHTMLElement> spFrameElement = spFrameBase;
		if (spFrameElement == NULL)
		{
			traceLog << "Can NOT get IHTMLElement in SearchFrameVisitor::GetFrameAttributeValue\n";
			return FALSE;
		}

		CComVariant vAttrValue;
		HRESULT     hRes = spFrameElement->getAttribute(CComBSTR(sAttributeName.c_str()), 0, &vAttrValue);

		if (hRes != S_OK)
		{
			traceLog << "Can not get the attribute " << sAttributeName << " in SearchFrameVisitor::GetFrameAttributeValue\n";
			return FALSE;
		}

		if ((vAttrValue.vt != VT_BSTR) || (NULL == vAttrValue.bstrVal))
		{
			// The attribute doesn't exist. Return empty string.
			*pAttributeValue = _T("");
		}
		else
		{
			USES_CONVERSION;
			*pAttributeValue = Common::TrimString(W2T(vAttrValue.bstrVal));
		}

		return TRUE;
	}


	/////////////////////////////////////////////////////////////////////////////////////////////////////
	SearchFrameCollectionVisitor::SearchFrameCollectionVisitor(SAFEARRAY* psa, LONG nSearchFlags) : 
			m_frameVisitor(psa)
	{
	}


	const list<CAdapt<CComQIPtr<IHTMLWindow2> > >& SearchFrameCollectionVisitor::GetFoundHtmlFrameCollection()
	{
		return m_foundFrameCollection;
	}


	BOOL SearchFrameCollectionVisitor::VisitFrame(CComQIPtr<IHTMLWindow2> spFrame)
	{
		BOOL bRes = m_frameVisitor.VisitFrame(spFrame);
		if (!bRes)
		{
			m_foundFrameCollection.push_back(spFrame);
		}

		return TRUE;
	}
}
