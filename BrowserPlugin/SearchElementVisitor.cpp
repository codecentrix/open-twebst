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
#include "SearchElementVisitor.h"
#include "HtmlHelpers.h"
using namespace std;
using namespace Common;


namespace FindElementAlgorithms
{
	// Constructor.
	SearchElementVisitor::SearchElementVisitor(SAFEARRAY* psa)
	{
		if (!Common::IsValidSafeArray(psa))
		{
			throw CreateInvalidParamException(_T("Invalid safearray parameter in SearchElementVisitor constructor"));
		}

		BOOL bRes = Common::GetDescriptorTokensList(psa, &m_tokens, &m_nSearchIndex);
		if (!bRes)
		{
			throw CreateInvalidParamException(_T("Can NOT get the list of descriptor tokens in SearchElementVisitor constructor"));
		}
	}


	SearchElementVisitor::~SearchElementVisitor(void)
	{
	}


	BOOL SearchElementVisitor::VisitElement(CComQIPtr<IHTMLElement> spElement)
	{
		ATLASSERT(spElement != NULL);

		for (list<DescriptorToken>::iterator itToken = m_tokens.begin(); itToken != m_tokens.end(); ++itToken)
		{
			// Get the value of the visited element for the current attribute.
			String sAttributeValue;
			BOOL   bRes = GetElementAttributeValue(spElement, itToken->m_sAttributeName, &sAttributeValue);
			if (FALSE == bRes)
			{
				traceLog << "Can NOT get the value of the current element " << itToken->m_sAttributeName << " in SearchElementVisitor::VisitElement no regexp case\n";
				return TRUE;	// Continue seraching.
			}

			BOOL bEqual;

			// Ignore last slash in SRC and HREF attributes.
			String sAttributePattern = itToken->m_sAttributePattern;
			if (!_tcsicmp(itToken->m_sAttributeName.c_str(), _T("src")) ||
				!_tcsicmp(itToken->m_sAttributeName.c_str(), _T("href")))
			{
				StripLastSlash(&sAttributePattern);
				StripLastSlash(&sAttributeValue);

				try
				{
					// Use a custom comparision method for src and href.
					bEqual = CompareHrefAndSrc(spElement, sAttributeValue, sAttributePattern);
				}
				catch (const ExceptionServices::Exception& except)
				{
					traceLog << except << "\n";
					return TRUE; // Continue searching.
				}

				if (!bEqual)
				{
					String sRelativeValue;
					bRes = GetElementAttributeValue(spElement, itToken->m_sAttributeName, &sRelativeValue, FALSE);
					if (FALSE == bRes)
					{
						traceLog << "Can NOT get the relative value of the current element " << itToken->m_sAttributeName << " in SearchElementVisitor::VisitElement no regexp case\n";
						return TRUE;	// Continue seraching.
					}

					StripLastSlash(&sRelativeValue);
					bEqual = Common::MatchWildcardPattern(sAttributePattern, sRelativeValue);
				}
			}
			else
			{
				bEqual = Common::MatchWildcardPattern(sAttributePattern, sAttributeValue);
			}

			if (DescriptorToken::TOKEN_MATCH == itToken->m_operator)
			{
				if (!bEqual)
				{
					return TRUE;
				}
			}
			else
			{
				ATLASSERT(DescriptorToken::TOKEN_NOT_MATCH == itToken->m_operator);
				if (bEqual)
				{
					return TRUE;
				}
			}
		}

		ATLASSERT(m_nSearchIndex >= 0);

		if (0 == m_nSearchIndex)
		{
			// Stop searching. Attributes and index matching.
			m_spFoundElement = spElement;
			return FALSE;
		}
		else
		{
			// Index does NOT match. Continue searching.
			m_nSearchIndex--;
			return TRUE;
		}
	}


	CComQIPtr<IHTMLElement> SearchElementVisitor::GetFoundHtmlElement()
	{
		return m_spFoundElement;
	}


	// TODO: <table> special case.
	BOOL SearchElementVisitor::GetElementAttributeValue(CComQIPtr<IHTMLElement> spElement,
	                                                    const String& sAttributeName,
                                                        String* pAttributeValue,
														BOOL    bFullAttribute)
	{
		ATLASSERT(spElement != NULL);
		ATLASSERT(pAttributeValue != NULL);

		if (!_tcsicmp(sAttributeName.c_str(), Common::TEXT_SEARCH_ATTRIBUTE))
		{
			USES_CONVERSION;
			*pAttributeValue = Common::TrimString(W2T(HtmlHelpers::GetTextAttributeValue(spElement)));
			return TRUE;
		}

		String sRealAttributeName; 

		if (!_tcsicmp(sAttributeName.c_str(), _T("class")))
		{
			// According to MSDN when using IHTMLElement::getAttribute to get the CLASS attribute,
			// className should be used as the attribute name.
			sRealAttributeName = _T("className");
		}
		else
		{
			sRealAttributeName = sAttributeName;
		}

		CComVariant vAttrValue;
		HRESULT hRes = spElement->getAttribute(CComBSTR(sRealAttributeName.c_str()), (bFullAttribute ? 0 : 2), &vAttrValue);
		if (hRes != S_OK)
		{
			traceLog << "Can not get the attribute" << sRealAttributeName << " in SearchElementVisitor::GetElementAttributeValue\n";
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


	BOOL SearchElementVisitor::CompareHrefAndSrc(IHTMLElement* pElement, const String& sElementUrl, const String& sCompareToUrl)
	{
		ATLASSERT(pElement != NULL);

		/////////////////////////////////////////////////////////////////////////////////////////////////
		// Step One: compare the links directly.
		if (sElementUrl == sCompareToUrl)
		{
			return TRUE;
		}


		/////////////////////////////////////////////////////////////////////////////////////////////////
		// Step Two: extract the file name from the element url and compare against the value coming from the script.

		// Get the document of the element.
		CComQIPtr<IDispatch> spDisp;
		HRESULT hRes = pElement->get_document(&spDisp);

		CComQIPtr<IHTMLDocument2> spDocument = spDisp;
		if (spDocument == NULL)
		{
			throw CreateException(_T("Can not get the document in HtmlHelpers::CompareHrefAndSrc"));
		}

		// Create a new <A> tag.
		CComQIPtr<IHTMLElement> spNewElement;
		hRes = spDocument->createElement(CComBSTR("a"), &spNewElement);
		if (FAILED(hRes))
		{
			throw CreateException(_T("Can not create a new element in HtmlHelpers::CompareHrefAndSrc"));
		}

		// Set the href attribute in the newly created anchor element.
		hRes = spNewElement->setAttribute(CComBSTR("href"), CComVariant(sElementUrl.c_str()), 0);
		if (FAILED(hRes))
		{
			throw CreateException(_T("Can not set the href in the html element in HtmlHelpers::CompareHrefAndSrc"));
		}

		CComQIPtr<IHTMLAnchorElement> spAnchor = spNewElement;
		if (spAnchor == NULL)
		{
			throw CreateException(_T("Query for IHTMLAnchorElement failed in HtmlHelpers::CompareHrefAndSrc"));
		}

		CComBSTR bstrFileName;
		hRes = spAnchor->get_nameProp(&bstrFileName);
		if (FAILED(hRes) || (bstrFileName == NULL))
		{
			throw CreateException(_T("IHTMLAnchorElement::get_nameProp failed in HtmlHelpers::CompareHrefAndSrc"));
		}

		USES_CONVERSION;
		if (Common::MatchWildcardPattern(W2T(bstrFileName), sCompareToUrl.c_str()))
		{
			// sCompareToUrl is the file name.
			return TRUE;
		}

		return FALSE;
	}




	////////////////////////////////////////////////////////////////////////////////////////////////////
	SearchElementCollectionVisitor::SearchElementCollectionVisitor (SAFEARRAY* psa) : m_elementVisitor(psa)
	{
	}


	const list<CAdapt<CComQIPtr<IHTMLElement> > >& SearchElementCollectionVisitor::GetFoundHtmlElementCollection() const
	{
		return m_foundElementCollection;
	}


	BOOL SearchElementCollectionVisitor::VisitElement(CComQIPtr<IHTMLElement> spElement)
	{
		BOOL bRes = m_elementVisitor.VisitElement(spElement);
		if (!bRes)
		{
			m_foundElementCollection.push_back(spElement);
		}

		return TRUE;
	}




	////////////////////////////////////////////////////////////////////////////////////////////////////
	SearchOptionIndexCollectionVisitor::SearchOptionIndexCollectionVisitor
	                                     (SAFEARRAY* psa) : 
		m_elementVisitor(psa), m_nCurrentCount(0)
	{
	}

	BOOL SearchOptionIndexCollectionVisitor::VisitElement(CComQIPtr<IHTMLElement> spElement)
	{
		BOOL bRes = m_elementVisitor.VisitElement(spElement);
		if (!bRes)
		{
			m_indexesList.push_back(m_nCurrentCount);
		}

		m_nCurrentCount++;
		return TRUE;
	}

	list<DWORD>& SearchOptionIndexCollectionVisitor::GetOptionIndexesList()
	{
		return m_indexesList;
	}


	////////////////////////////////////////////////////////////////////////////////////////////////////
	// SearchSelectedOptionVisitor
	CComQIPtr<IHTMLElement> SearchSelectedOptionVisitor::GetFoundHtmlElement()
	{
		return m_spFoundElement;
	}

	BOOL SearchSelectedOptionVisitor::VisitElement(CComQIPtr<IHTMLElement> spElement)
	{
		ATLASSERT(spElement != NULL);

		CComQIPtr<IHTMLOptionElement> spOptionElem = spElement;
		if (spOptionElem != NULL)
		{
			VARIANT_BOOL vbIsSelected; 
			HRESULT      hRes = spOptionElem->get_selected(&vbIsSelected);

			if (SUCCEEDED(hRes) && (VARIANT_TRUE == vbIsSelected))
			{
				// Selected option found. Stop searching.
				m_spFoundElement = spElement;
				return FALSE;
			}
		}

		return TRUE; // Continue seraching.
	}


	////////////////////////////////////////////////////////////////////////////////////////////////////
	// SearchAllSelectedOptionsVisitor
	const list<CAdapt<CComQIPtr<IHTMLElement> > >& SearchAllSelectedOptionsVisitor::GetFoundHtmlElementCollection() const
	{
		return m_foundElementCollection;
	}


	BOOL SearchAllSelectedOptionsVisitor::VisitElement(CComQIPtr<IHTMLElement> spElement)
	{
		BOOL bRes = m_elementVisitor.VisitElement(spElement);
		if (!bRes)
		{
			m_foundElementCollection.push_back(spElement);
		}

		return TRUE;
	}
}
