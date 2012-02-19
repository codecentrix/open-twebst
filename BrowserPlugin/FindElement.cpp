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
#include "HtmlHelpers.h"
#include "Visitor.h"
#include "SearchElementVisitor.h"
#include "FindElement.h"
#include "SearchCondition.h"

// Local function declarations.
namespace FindElementAlgorithms
{
	BOOL VisitParentElementCollection  (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);
	BOOL VisitElementCollection        (CComQIPtr<IWebBrowser2> spBrowser, const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);
	BOOL VisitElementCollection        (CComQIPtr<IHTMLWindow2> spWindow,  const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);
	BOOL VisitChildrenElementCollection(CComQIPtr<IHTMLWindow2> spWindow,  const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);
	BOOL VisitElementCollection        (CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);
	BOOL VisitChildrenElementCollection(CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, Visitor* pVisitor);

	CComQIPtr<IHTMLElementCollection> GetElementCollectionByTag(CComQIPtr<IHTMLElementCollection> spCollection, const CComBSTR& bstrTagName);
	BOOL GetTagAndInputType(const CComBSTR& bstrCompoundTagName, CComBSTR& bstrTag, CComBSTR& bstrType);
	BOOL IsValidInputType  (const String& sInputType);
	CComQIPtr<ILocalElementCollection> CreateElementCollection(const list<CAdapt<CComQIPtr<IHTMLElement> > >& elementList);
}


list<DWORD> FindElementAlgorithms::FindOptionsIndexes(CComQIPtr<IHTMLElement> spSelectElement, BSTR bstrCondition)
{
	ATLASSERT(spSelectElement != NULL);
	ATLASSERT(bstrCondition   != NULL);

	USES_CONVERSION;
	String sCondition = _T("text=");
	sCondition += W2T(bstrCondition);

	SearchCondition sc = sCondition.c_str();
	SearchOptionIndexCollectionVisitor visitor(sc);
	VisitElementCollection(spSelectElement, CComBSTR("option"), &visitor);

	list<DWORD> result;
	result.swap(visitor.GetOptionIndexesList());
	return result;
}


// Returns NULL if the input list is empty.
CComQIPtr<ILocalElementCollection> FindElementAlgorithms::CreateElementCollection(const list<CAdapt<CComQIPtr<IHTMLElement> > >& elementList)
{
	// Create an ILocalElementCollection object.
	ILocalElementCollection* pNewElementCollection = NULL;
	HRESULT hRes = CComCoClass<CLocalElementCollection>::CreateInstance(&pNewElementCollection);

	if (pNewElementCollection != NULL)
	{
		ATLASSERT(SUCCEEDED(hRes));
		CLocalElementCollection* pElementCollectionObject = static_cast<CLocalElementCollection*>(pNewElementCollection); // Down cast !!!
		ATLASSERT(pElementCollectionObject != NULL);
		pElementCollectionObject->InitCollection(elementList);

		CComQIPtr<ILocalElementCollection> spResult;
		spResult.Attach(pNewElementCollection);
		return spResult;
	}
	else
	{
		throw CreateException(_T("Can not create an LocalElementCollection object in FindElementAlgorithms::CreateElementCollection"));
	}
}


CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindAllElements
                                   (CComQIPtr<IHTMLElement> spElement,
                                    const CComBSTR& bstrCompoundTagName,
                                    SAFEARRAY* psa)
{
	ATLASSERT(spElement           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementCollectionVisitor visitor(psa);
	VisitElementCollection(spElement, bstrCompoundTagName, &visitor);

	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}


CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindChildrenElements
                                   (CComQIPtr<IHTMLElement> spElement,
                                    const CComBSTR& bstrCompoundTagName,
                                    SAFEARRAY* psa)
{
	ATLASSERT(spElement           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementCollectionVisitor visitor(psa);
	VisitChildrenElementCollection(spElement, bstrCompoundTagName, &visitor);

	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}


CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindAllElements
                                   (CComQIPtr<IHTMLWindow2> spWindow,
                                    const CComBSTR& bstrCompoundTagName,
                                    SAFEARRAY* psa)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementCollectionVisitor visitor(psa);
	VisitElementCollection(spWindow, bstrCompoundTagName, &visitor);

	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}


CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindChildrenElements
                                   (CComQIPtr<IHTMLWindow2> spWindow,
                                    const CComBSTR& bstrCompoundTagName,
                                    SAFEARRAY* psa)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementCollectionVisitor visitor(psa);
	VisitChildrenElementCollection(spWindow, bstrCompoundTagName, &visitor);

	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}


// Public functions definitions.
CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindAllElements
                                                      (CComQIPtr<IWebBrowser2> spBrowser,
                                                       const CComBSTR& bstrCompoundTagName,
                                                       SAFEARRAY* psa)
{
	ATLASSERT(spBrowser           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementCollectionVisitor visitor(psa);
	VisitElementCollection(spBrowser, bstrCompoundTagName, &visitor);

	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindElement(CComQIPtr<IWebBrowser2> spBrowser,
                                                           const CComBSTR& bstrCompoundTagName,
                                                           SAFEARRAY* psa)
{
	ATLASSERT(spBrowser           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitElementCollection(spBrowser, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindElement(CComQIPtr<IHTMLWindow2> spWindow,
                                                           const CComBSTR& bstrCompoundTagName,
                                                           SAFEARRAY* psa)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitElementCollection(spWindow, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindChildElement(CComQIPtr<IHTMLWindow2> spWindow,
                                                                const CComBSTR& bstrCompoundTagName,
                                                                SAFEARRAY* psa)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitChildrenElementCollection(spWindow, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindElement(CComQIPtr<IHTMLElement> spElement,
														   const CComBSTR& bstrCompoundTagName,
														   SAFEARRAY* psa)
{
	ATLASSERT(spElement           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitElementCollection(spElement, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindParentElement
	(CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, SAFEARRAY* psa)
{
	ATLASSERT(spElement           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitParentElementCollection(spElement, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindChildElement(CComQIPtr<IHTMLElement> spElement,
														        const CComBSTR& bstrCompoundTagName,
														        SAFEARRAY* psa)
{
	ATLASSERT(spElement           != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);
	ATLASSERT(psa                 != NULL);

	SearchElementVisitor visitor(psa);
	VisitChildrenElementCollection(spElement, bstrCompoundTagName, &visitor);

	return visitor.GetFoundHtmlElement();
}


CComQIPtr<IHTMLElement> FindElementAlgorithms::FindSelectedOption(CComQIPtr<IHTMLSelectElement> spSelectElement)
{
	ATLASSERT(spSelectElement != NULL);

	SearchSelectedOptionVisitor visitor;
	CComQIPtr<IHTMLElement>     spElement = spSelectElement;

	if (spElement == NULL)
	{
		throw CreateException(_T("Cannot query for IHTMLElement in FindElementAlgorithms::FindSelectedOption"));
	}

	VisitElementCollection(spElement, CComBSTR("option"), &visitor);
	return visitor.GetFoundHtmlElement();
}


CComQIPtr<ILocalElementCollection> FindElementAlgorithms::FindAllSelectedOptions(CComQIPtr<IHTMLSelectElement> spSelectElement)
{
	ATLASSERT(spSelectElement != NULL);

	SearchAllSelectedOptionsVisitor visitor;
	CComQIPtr<IHTMLElement>         spElement = spSelectElement;

	if (spElement == NULL)
	{
		throw CreateException(_T("Cannot query for IHTMLElement in FindElementAlgorithms::FindAllSelectedOptions"));
	}

	VisitElementCollection(spElement, CComBSTR("option"), &visitor);
	return CreateElementCollection(visitor.GetFoundHtmlElementCollection());
}



///////////////////////////////////////////////////////////////////////////////////////////////////////////

// The function browses the direct children collection of spElement.
BOOL FindElementAlgorithms::VisitChildrenElementCollection(CComQIPtr<IHTMLElement> spElement,
                                                           const CComBSTR& bstrCompoundTagName,
                                                           Visitor* pVisitor)
{
	ATLASSERT(spElement != NULL);
	ATLASSERT(pVisitor  != NULL);
	ATLASSERT(bstrCompoundTagName.Length() != 0);

	// Get the collection of direct children elements of the spElement.
	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = spElement->get_children(&spDisp);
	if ((hRes != S_OK) || (spDisp == NULL))
	{
		throw CreateException(_T("IHTMLElement::get_children failed in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	CComQIPtr<IHTMLElementCollection> spChildrenCollection = spDisp;
	if (spChildrenCollection == NULL)
	{
		throw CreateException(_T("Query for interface IHTMLElementCollection failed in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	CComBSTR bstrTagName;
	CComBSTR bstrInputType;
	BOOL     bIsInputType = GetTagAndInputType(bstrCompoundTagName, bstrTagName, bstrInputType);

	// From the direct children collection keep only the elements with the given tag name.
	CComQIPtr<IHTMLElementCollection> spTagCollection = GetElementCollectionByTag(spChildrenCollection, bstrTagName);
	spChildrenCollection.Release();

	// Get the size of the tag collection.
	long nCollectionSize;
	hRes = spTagCollection->get_length(&nCollectionSize);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLElementCollection::get_length failed in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		CComQIPtr<IDispatch> spItemDisp;
		hRes = spTagCollection->item(CComVariant(i), CComVariant(), &spItemDisp);
		if ((hRes != S_OK) || (spItemDisp == NULL))
		{
			traceLog << "Can not get the" << i << "th item in FindElementAlgorithms::VisitChildrenElementCollection\n";
			continue;
		}

		CComQIPtr<IHTMLElement> spCurrentElement = spItemDisp;
		if (spCurrentElement == NULL)
		{
			traceLog << "Query for IHTMLElement from IDispatch failed for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
			continue;
		}

		if (bIsInputType)
		{
			// Get the input element object.
			CComQIPtr<IHTMLInputElement> spInputElement = spCurrentElement;
			if (spInputElement == NULL)
			{
				traceLog << "Query for IHTMLInputElement from IHTMLElement failed for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
				continue;
			}

			// Get the input type of the current input element.
			CComBSTR bstrCurrentType;
			hRes = spInputElement->get_type(&bstrCurrentType);
			if ((hRes != S_OK) || (NULL == bstrCurrentType.m_str))
			{
				traceLog << "Can not get the input type for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
				continue;
			}

			if (_wcsicmp(bstrInputType, bstrCurrentType))
			{
				// The input element is not of the type to be found.
				continue;
			}
		}

		// Visit the current html element.
		BOOL bRes = pVisitor->VisitElement(spCurrentElement);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


// This function  browses all children of spElement. Return TRUE to continue browsing or FALSE to stop.
BOOL FindElementAlgorithms::VisitElementCollection(CComQIPtr<IHTMLElement> spElement,
											       const CComBSTR& bstrCompoundTagName,
                                                   Visitor* pVisitor)
{
	ATLASSERT(spElement != NULL);
	ATLASSERT(pVisitor  != NULL);
	ATLASSERT(bstrCompoundTagName.Length() != 0);

	// Get the collection of all elements in the scope of spElement.
	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = spElement->get_all(&spDisp);
	if ((hRes != S_OK) || (spDisp == NULL))
	{
		throw CreateException(_T("IHTMLElement::get_all failed in FindElementAlgorithms::VisitElementCollection"));
	}

	CComQIPtr<IHTMLElementCollection> spAllCollection = spDisp;
	if (spAllCollection == NULL)
	{
		throw CreateException(_T("Query for interface IHTMLElementCollection failed in FindElementAlgorithms::VisitElementCollection"));
	}

	CComBSTR bstrTagName;
	CComBSTR bstrInputType;
	BOOL     bIsInputType = GetTagAndInputType(bstrCompoundTagName, bstrTagName, bstrInputType);

	// From the "all" collection keep only the elements with the given tag name.
	CComQIPtr<IHTMLElementCollection> spTagCollection = GetElementCollectionByTag(spAllCollection, bstrTagName);
	spAllCollection.Release();

	// Get the size of the tag collection.
	long nCollectionSize;
	hRes = spTagCollection->get_length(&nCollectionSize);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLElementCollection::get_length failed in FindElementAlgorithms::VisitElementCollection"));
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		CComQIPtr<IDispatch> spItemDisp;
		hRes = spTagCollection->item(CComVariant(i), CComVariant(), &spItemDisp);
		if ((hRes != S_OK) || (spItemDisp == NULL))
		{
			traceLog << "Can not get the" << i << "th item in FindElementAlgorithms::VisitElementCollection\n";
			continue;
		}

		CComQIPtr<IHTMLElement> spCurrentElement = spItemDisp;
		if (spCurrentElement == NULL)
		{
			traceLog << "Query IHTMLElement from IDispatch failed for index" << i << " in FindElementAlgorithms::VisitElementCollection\n";
			continue;
		}

		if (bIsInputType)
		{
			// Get the input element object.
			CComQIPtr<IHTMLInputElement> spInputElement = spCurrentElement;
			if (spInputElement == NULL)
			{
				traceLog << "Query for IHTMLInputElement from IHTMLElement failed for index" << i << " in FindElementAlgorithms::VisitElementCollection\n";
				continue;
			}

			// Get the input type of the current input element.
			CComBSTR bstrCurrentType;
			hRes = spInputElement->get_type(&bstrCurrentType);
			if ((hRes != S_OK) || (NULL == bstrCurrentType.m_str))
			{
				traceLog << "Can not get the input type for index" << i << " in FindElementAlgorithms::VisitElementCollection\n";
				continue;
			}

			if (_wcsicmp(bstrInputType, bstrCurrentType))
			{
				// The input element is not of the type to be found.
				continue;
			}
		}

		// Visit the current html element.
		BOOL bRes = pVisitor->VisitElement(spCurrentElement);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


// This function  browses all children of spElement. Return TRUE to continue browsing or FALSE to stop.
BOOL FindElementAlgorithms::VisitParentElementCollection
	(CComQIPtr<IHTMLElement> spElement, const CComBSTR& bstrCompoundTagName, Visitor* pVisitor)
{
	ATLASSERT(spElement != NULL);
	ATLASSERT(pVisitor  != NULL);
	ATLASSERT(bstrCompoundTagName.m_str != NULL);


	CComBSTR bstrTagName;
	CComBSTR bstrInputType;
	BOOL     bIsInputType = GetTagAndInputType(bstrCompoundTagName, bstrTagName, bstrInputType);

	// Browse the parent collection.
	CComQIPtr<IHTMLElement> spCurrentElement = spElement;

	while (TRUE)
	{

		CComQIPtr<IHTMLElement> spParentElement;
		HRESULT hRes = spCurrentElement->get_parentElement(&spParentElement);

		if (FAILED(hRes) || (spParentElement == NULL))
		{
			// Reach the end of the hierarchy or error.
			break;
		}

		spCurrentElement = spParentElement;

		// If no tag name is provided as argument of FindParentElement then search any element.
		if (bstrCompoundTagName.Length() > 0)
		{
			CComBSTR bstrCrntTagName;
			hRes = spCurrentElement->get_tagName(&bstrCrntTagName);

			if (FAILED(hRes) || (bstrCrntTagName == NULL))
			{
				traceLog << "get_tagName failed in FindElementAlgorithms::VisitParentElementCollection\n";
				break;
			}

			if (_wcsicmp(bstrTagName, bstrCrntTagName))
			{
				// Tag name doesn't match.
				continue;
			}

			if (bIsInputType)
			{
				// Get the input element object.
				CComQIPtr<IHTMLInputElement> spInputElement = spCurrentElement;
				if (spInputElement == NULL)
				{
					traceLog << "Query for IHTMLInputElement from IHTMLElement failed in FindElementAlgorithms::VisitParentElementCollection\n";
					continue;
				}

				// Get the input type of the current input element.
				CComBSTR bstrCurrentType;
				hRes = spInputElement->get_type(&bstrCurrentType);
				if ((hRes != S_OK) || (NULL == bstrCurrentType.m_str))
				{
					traceLog << "Can not get the input type in FindElementAlgorithms::VisitParentElementCollection\n";
					continue;
				}

				if (_wcsicmp(bstrInputType, bstrCurrentType))
				{
					// The input element is not of the type to be found.
					continue;
				}
			}
		}

		// Visit the current html element.
		BOOL bRes = pVisitor->VisitElement(spCurrentElement);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


CComQIPtr<IHTMLElementCollection> FindElementAlgorithms::GetElementCollectionByTag(CComQIPtr<IHTMLElementCollection> spCollection,
                                                                                   const CComBSTR& bstrTagName)
{
	ATLASSERT(spCollection != NULL);
	ATLASSERT(bstrTagName  != NULL);

	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = spCollection->tags(CComVariant(bstrTagName), &spDisp);
	if ((hRes != S_OK) || (spDisp == NULL))
	{
		throw CreateException(_T("IHTMLElementCollection::tags failed in FindElementAlgorithms::GetElementCollectionByTag"));
	}

	CComQIPtr<IHTMLElementCollection> spResultCollection = spDisp;
	if (spResultCollection == NULL)
	{
		throw CreateException(_T("Query for interface IHTMLElementCollection failed in FindElementAlgorithms::GetElementCollectionByTag"));
	}

	return spResultCollection;
}


// Visit the html elements inside a IHTMLWindow2 frame. This function does not search recurrently in the subframes of the spWindow frame.
BOOL FindElementAlgorithms::VisitChildrenElementCollection(CComQIPtr<IHTMLWindow2> spWindow,
                                                           const CComBSTR& bstrCompoundTagName,
                                                           Visitor* pVisitor)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(pVisitor            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);


	// Get the html document of the html window.
	CComQIPtr<IHTMLDocument2> spDoc = HtmlHelpers::HtmlWindowToHtmlDocument(spWindow);
	if (spDoc == NULL)
	{
		throw CreateException(_T("Can not get the IHTMLDocument2 from IHTMLWindow2 in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	// Get the "all" collection of the html document.
	CComQIPtr<IHTMLElementCollection> spAllCollection;
	HRESULT hRes = spDoc->get_all(&spAllCollection);
	if ((hRes != S_OK) || (spAllCollection == NULL))
	{
		throw CreateException(_T("Can not get the all collection of the document in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	CComBSTR bstrTagName;
	CComBSTR bstrInputType;
	BOOL     bIsInputType = GetTagAndInputType(bstrCompoundTagName, bstrTagName, bstrInputType);

	// From the "all" collection keep only the elements with the given tag name.
	CComQIPtr<IHTMLElementCollection> spTagCollection = GetElementCollectionByTag(spAllCollection, bstrTagName);
	spAllCollection.Release();

	// Get the size of the tag collection.
	long nCollectionSize;
	hRes = spTagCollection->get_length(&nCollectionSize);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLElementCollection::get_length failed in FindElementAlgorithms::VisitChildrenElementCollection"));
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		CComQIPtr<IDispatch> spItemDisp;
		hRes = spTagCollection->item(CComVariant(i), CComVariant(), &spItemDisp);
		if ((hRes != S_OK) || (spItemDisp == NULL))
		{
			traceLog << "Can not get the" << i << "th item in FindElementAlgorithms::VisitChildrenElementCollection\n";
			continue;
		}

		CComQIPtr<IHTMLElement> spCurrentElement = spItemDisp;
		if (spCurrentElement == NULL)
		{
			traceLog << "Query IHTMLElement from IDispatch failed for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
			continue;
		}

		if (bIsInputType)
		{
			// Get the input element object.
			CComQIPtr<IHTMLInputElement> spInputElement = spCurrentElement;
			if (spInputElement == NULL)
			{
				traceLog << "Query for IHTMLInputElement from IHTMLElement failed for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
				continue;
			}

			// Get the input type of the current input element.
			CComBSTR bstrCurrentType;
			hRes = spInputElement->get_type(&bstrCurrentType);
			if ((hRes != S_OK) || (NULL == bstrCurrentType.m_str))
			{
				traceLog << "Can not get the input type for index" << i << " in FindElementAlgorithms::VisitChildrenElementCollection\n";
				continue;
			}

			if (_wcsicmp(bstrInputType, bstrCurrentType))
			{
				// The input element is not of the type to be found.
				continue;
			}
		}

		// Visit the current html element.
		BOOL bRes = pVisitor->VisitElement(spCurrentElement);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


// Recurently browse the html elements in the spWindow frame and subframes.
BOOL FindElementAlgorithms::VisitElementCollection(CComQIPtr<IHTMLWindow2> spWindow,
                                                   const CComBSTR& bstrCompoundTagName,
                                                   Visitor* pVisitor)
{
	ATLASSERT(spWindow            != NULL);
	ATLASSERT(pVisitor            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);

	// Visit the html elements in spWindow frame.
	BOOL bRes = VisitChildrenElementCollection(spWindow, bstrCompoundTagName, pVisitor);
	if (!bRes)
	{
		// Stop browsing.
		return FALSE;
	}

	// Get the subframes collection of the spWindow html window.
	CComQIPtr<IHTMLFramesCollection2> spFrameCollection;
	HRESULT hRes = spWindow->get_frames(&spFrameCollection);
	if ((hRes != S_OK) || (spFrameCollection == NULL))
	{
		throw CreateException(_T("IHTMLWindow2::get_frames failed in FindElementAlgorithms::VisitElementCollection"));
	}

	// Get the number of frames in the spFrameCollection.
	long nNumberOfSubframes = 0;
	hRes = spFrameCollection->get_length(&nNumberOfSubframes);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLFramesCollection2::get_length failed in FindElementAlgorithms::VisitElementCollection"));
	}

	// Browse the subframes collection.
	for (long i = 0; i < nNumberOfSubframes; ++i)
	{
		CComVariant		varSubframe;
		CComVariant		varIndex(i);
		hRes = spFrameCollection->item(&varIndex, &varSubframe);
		if (hRes!= S_OK)
		{
			traceLog << "IHTMLFramesCollection2::item failed for index " << i << " in FindElementAlgorithms::VisitElementCollection\n";
			continue;
		}

		if ((varSubframe.vt != VT_DISPATCH) || (NULL == varSubframe.pdispVal))
		{
			traceLog << "IHTMLFramesCollection2::item returned invalid variant for index " << i << " in FindElementAlgorithms::VisitElementCollection\n";
			continue;
		}

		// Get the current subframe;
		CComQIPtr<IHTMLWindow2>	spCurrentFrame = varSubframe.pdispVal;
		if (spCurrentFrame == NULL)
		{
			traceLog << "Query for IHTMLWindow2 failed for index " << i << " in FindElementAlgorithms::VisitElementCollection\n";
			continue;
		}

		// Recurrently visit the subframes.
		bRes = VisitElementCollection(spCurrentFrame, bstrCompoundTagName, pVisitor);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


BOOL FindElementAlgorithms::VisitElementCollection(CComQIPtr<IWebBrowser2> spBrowser,
                                                   const CComBSTR& bstrCompoundTagName,
                                                   Visitor* pVisitor)
{
	ATLASSERT(spBrowser           != NULL);
	ATLASSERT(pVisitor            != NULL);
	ATLASSERT(bstrCompoundTagName != NULL);

	// Get a IHTMLWindow2 object from IWebBrowser2 object.
	CComQIPtr<IHTMLWindow2> spMainFrame = HtmlHelpers::HtmlWebBrowserToHtmlWindow(spBrowser);
	if (spMainFrame != NULL)
	{
		return VisitElementCollection(spMainFrame, bstrCompoundTagName, pVisitor);
	}
	else
	{
		throw CreateException(_T("Can not get the IHTMLWindow2 from IWebBrowser2 in FindElementAlgorithms::VisitElementCollection"));
	}
}


// The bstrCompoundTagName is the tag name. For the <input> tag, bstrCompoundTagName is "INPUT TYPE"
// Example: INPUT TEXT.
// Returns TRUE if it is an <INPUT> element, FALSE otherwise.
BOOL FindElementAlgorithms::GetTagAndInputType(const CComBSTR& bstrCompoundTagName, CComBSTR& bstrTag, CComBSTR& bstrType)
{
	ATLASSERT(bstrCompoundTagName != NULL);

	USES_CONVERSION;
	String sCompoundTagName = Common::TrimString(W2T(bstrCompoundTagName));

	// Find the first space in bstrCompoundTagName.
	const TCHAR* pFirstSpace = _tcschr(sCompoundTagName.c_str(), _T(' '));
	if (NULL == pFirstSpace)
	{
		// It is not an input tag.
		bstrTag  = sCompoundTagName.c_str();
		bstrType = L"";
		return FALSE;
	}

	const WCHAR INPUT_TAG[]  = L"INPUT";
	const int   nInputTagLen = sizeof(INPUT_TAG) / sizeof(INPUT_TAG[0]) - 1;
	if (!_wcsnicmp(bstrCompoundTagName, INPUT_TAG, nInputTagLen))
	{
		String sInputType = Common::TrimString(pFirstSpace);
		if (sInputType.empty())
		{
			bstrTag  = INPUT_TAG;
			bstrType = L"";
			return FALSE;
		}
		else if (IsValidInputType(sInputType))
		{
			bstrTag  = INPUT_TAG;
			bstrType = T2W(const_cast<LPTSTR>(sInputType.c_str()));
			return TRUE;
		}
		else
		{
			throw CreateInvalidParamException(_T("Invalid input type in FindElementAlgorithms::GetTagName"));
		}
	}
	else
	{
		throw CreateInvalidParamException(_T("Invalid tag in FindElementAlgorithms::GetTagName"));
	}
}


BOOL FindElementAlgorithms::IsValidInputType(const String& sInputType)
{
	const TCHAR* const VALID_INPUT_TYPES[] = 
	{
		_T("button"), _T("checkbox"), _T("file"), _T("hidden"), _T("image"), _T("password"), _T("radio"),
		_T("reset"),  _T("submit"),   _T("text")
	};
	const int NUMBER_OF_VALID_INPUT_TYPES = sizeof(VALID_INPUT_TYPES) / sizeof(VALID_INPUT_TYPES[0]);

	for (int i = 0; i < NUMBER_OF_VALID_INPUT_TYPES; ++i)
	{
		if (!_tcsicmp(sInputType.c_str(), VALID_INPUT_TYPES[i]))
		{
			return TRUE;
		}
	}

	return FALSE;
}
