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
#include "Common.h"
#include "Visitor.h"
#include "FindElement.h"

namespace FindElementAlgorithms
{
	//////////////////////////////////////////////////////////////////////////////////////////////////
	class SearchElementVisitor : public FindElementAlgorithms::Visitor
	{
	public:
		SearchElementVisitor(SAFEARRAY* psa);
		~SearchElementVisitor(void);
		virtual BOOL VisitElement(CComQIPtr<IHTMLElement> spElement);
		CComQIPtr<IHTMLElement> GetFoundHtmlElement();

	private:
		BOOL GetElementAttributeValue(CComQIPtr<IHTMLElement> spElement, const String& sAttributeName, String* pAttributeValue, BOOL bFullAttribute = TRUE);
		BOOL CompareHrefAndSrc(IHTMLElement* pElement, const String& sElementCompleteUrl, const String& sCompareToUrl);

	private:
		CComQIPtr<IHTMLElement>    m_spFoundElement;
		std::list<DescriptorToken> m_tokens;
		int                        m_nSearchIndex;
	};


	//////////////////////////////////////////////////////////////////////////////////////////////////
	class SearchElementCollectionVisitor : public FindElementAlgorithms::Visitor
	{
	public:
		SearchElementCollectionVisitor (SAFEARRAY* psa);
		virtual BOOL VisitElement      (CComQIPtr<IHTMLElement> spElement);
		const list<CAdapt<CComQIPtr<IHTMLElement> > >& GetFoundHtmlElementCollection() const;

	private:
		SearchElementVisitor                    m_elementVisitor;
		list<CAdapt<CComQIPtr<IHTMLElement> > > m_foundElementCollection;
	};


	//////////////////////////////////////////////////////////////////////////////////////////////////
	class SearchOptionIndexCollectionVisitor : public FindElementAlgorithms::Visitor
	{
	public:
		SearchOptionIndexCollectionVisitor (SAFEARRAY* psa);
		virtual BOOL VisitElement      (CComQIPtr<IHTMLElement> spElement);
		list<DWORD>& GetOptionIndexesList();

	private:
		int                  m_nCurrentCount;
		SearchElementVisitor m_elementVisitor;
		list<DWORD>          m_indexesList;
	};


	//////////////////////////////////////////////////////////////////////////////////////////////////
	class SearchSelectedOptionVisitor : public FindElementAlgorithms::Visitor
	{
	public:
		virtual BOOL VisitElement(CComQIPtr<IHTMLElement> spElement);
		CComQIPtr<IHTMLElement> GetFoundHtmlElement();

	private:
		CComQIPtr<IHTMLElement> m_spFoundElement;
	};

	//////////////////////////////////////////////////////////////////////////////////////////////////
	class SearchAllSelectedOptionsVisitor : public FindElementAlgorithms::Visitor
	{
	public:
		virtual BOOL VisitElement(CComQIPtr<IHTMLElement> spElement);
		const list<CAdapt<CComQIPtr<IHTMLElement> > >& GetFoundHtmlElementCollection() const;

	private:
		SearchSelectedOptionVisitor             m_elementVisitor;
		list<CAdapt<CComQIPtr<IHTMLElement> > > m_foundElementCollection;
	};
}
