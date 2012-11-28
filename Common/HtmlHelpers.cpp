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
#include "Common.h"

namespace HtmlHelpers
{
	CComBSTR GetAccessibleNameUsingDom(CComQIPtr<IHTMLElement> spHtmlElement);
	CComBSTR GetRightAccessibleName(CComQIPtr<IHTMLElement> spHtmlElement);
	CComBSTR GetLeftAccessibleName(CComQIPtr<IHTMLElement> spHtmlElement);
	CComBSTR GetRightText(CComQIPtr<IHTMLDOMNode> spNode);
	CComBSTR GetLeftText(CComQIPtr<IHTMLDOMNode> spNode);
	CComQIPtr<IHTMLMapElement> GetAreaMap(CComQIPtr<IHTMLAreaElement> spArea);
	CComQIPtr<IHTMLImgElement> GetMapImage(CComQIPtr<IHTMLMapElement> spMap);
	BOOL     IsValidForComputingTextUsingAA(CComQIPtr<IHTMLElement> spHtmlElement);
	BOOL     IsValidForUsingInnerText(CComQIPtr<IHTMLElement> spHtmlElement);
	BOOL     NeedFireOnclick(IHTMLElement* pElement);
	HRESULT  RightClickElementAt(IHTMLElement* pElement, LONG x, LONG y);
	
	CComQIPtr<IHTMLElement> FindFrameElementByTickCount(CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCrntTime);
	CComQIPtr<IHTMLElementCollection> GetElementCollectionByTag(CComQIPtr<IHTMLElementCollection> spCollection, const CComBSTR& bstrTagName);
	CComQIPtr<IHTMLElement> FindFrameElementByTickCount(CComQIPtr<IHTMLElementCollection> spElemCollection, const CComBSTR& bstrCrntTime);
}


// Converts a IHTMLWindow2 object to a IHTMLDocument2. Returns NULL in case of failure.
CComQIPtr<IHTMLDocument2> HtmlHelpers::HtmlWindowToHtmlDocument(CComQIPtr<IHTMLWindow2> spWindow)
{
	ATLASSERT(spWindow != NULL);

	CComQIPtr<IHTMLDocument2> spDocument;
	HRESULT hRes = spWindow->get_document(&spDocument);
	
	if ((S_OK == hRes) && (spDocument != NULL))
	{
		// The html document was properly retrieved.
		return spDocument;
	}

	// hRes could be E_ACCESSDENIED that means a security restriction that
	// prevents scripting across frames that loads documents from different internet domains.
	traceLog << "IHTMLWindow2::get_document failed with code:" << hRes << "\n";

	CComQIPtr<IWebBrowser2>	spBrws = HtmlWindowToHtmlWebBrowser(spWindow);
	if (spBrws == NULL)
	{
		traceLog << "Can not get the IWebBrowser2 object in HtmlHelpers::HtmlWindowToHtmlDocument " << hRes << "\n";
		return CComQIPtr<IHTMLDocument2>();
	}

	// Get the document object from the IWebBrowser2 object.
	CComQIPtr<IDispatch> spDisp;
	hRes = spBrws->get_Document(&spDisp);
	spDocument = spDisp;

	return spDocument;
}


// Converts a IHTMLWindow2 object to a IWebBrowser2. Returns NULL in case of failure.
CComQIPtr<IWebBrowser2> HtmlHelpers::HtmlWindowToHtmlWebBrowser(CComQIPtr<IHTMLWindow2> spWindow)
{
	ATLASSERT(spWindow != NULL);

	CComQIPtr<IServiceProvider>	spServiceProvider = spWindow;
	if (spServiceProvider == NULL)
	{
		traceLog << "Can not get IServiceProvider in HtmlHelpers::HtmlWindowToHtmlWebBrowser\n";
		return CComQIPtr<IWebBrowser2>();
	}

	CComQIPtr<IWebBrowser2> spWebBrws;
	HRESULT hRes = spServiceProvider->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**)&spWebBrws);
	if (hRes != S_OK)
	{
		traceLog << "Can not get IWebBrowser2 in HtmlHelpers::HtmlWindowToHtmlWebBrowser\n";
		return CComQIPtr<IWebBrowser2>();
	}

	return spWebBrws;
}


// Converts a IWebBrowser2 object to a IHTMLWindow2. Returns NULL in case of failure.
CComQIPtr<IHTMLWindow2> HtmlHelpers::HtmlWebBrowserToHtmlWindow(CComQIPtr<IWebBrowser2> spBrowser)
{
	// The window found by code bellow behaves weird when get_top or get_parent is called on it.
	// If it is a top window, top and parent should return the same object but it doesn't.
	/*CComQIPtr<IServiceProvider>	spServiceProvider = spBrowser;
	if (spServiceProvider == NULL)
	{
		traceLog << "Can not get IServiceProvider in HtmlHelpers::HtmlWebBrowserToHtmlWindow\n";
		return CComQIPtr<IHTMLWindow2>();
	}

	CComQIPtr<IHTMLWindow2> spWindow;
	HRESULT hRes = spServiceProvider->QueryService(IID_IHTMLWindow2, IID_IHTMLWindow2, (void**)&spWindow);
	if (hRes != S_OK)
	{
		traceLog << "Can not get IHTMLWindow2 in HtmlHelpers::HtmlWebBrowserToHtmlWindow\n";
		return CComQIPtr<IHTMLWindow2>();
	}

	return spWindow;*/

	ATLASSERT(spBrowser != NULL);
	CComQIPtr<IHTMLWindow2> spWindow;

	// Get the document of the browser.
	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = spBrowser->get_Document(&spDisp);
	if (FAILED(hRes) || (spDisp == NULL))
	{
		traceLog << "IWebBrowser2::get_Document failed in HtmlHelpers::HtmlWebBrowserToHtmlWindow\n";
		return CComQIPtr<IHTMLWindow2>();
	}

	// Get the window of the document.
	CComQIPtr<IHTMLDocument2> spDoc = spDisp;
	if (spDoc != NULL)
	{
		hRes = spDoc->get_parentWindow(&spWindow);
		if (FAILED(hRes) || (spWindow == NULL))
		{
			traceLog << "IHTMLDocument2::get_parentWindow failed in HtmlHelpers::HtmlWebBrowserToHtmlWindow\n";
			return CComQIPtr<IHTMLWindow2>();
		}
	}
	else
	{
		traceLog << "Can NOT query for IHTMLDocument2 in HtmlHelpers::HtmlWebBrowserToHtmlWindow\n";
	}

	return spWindow;
}


CComBSTR HtmlHelpers::GetAccessibleNameUsingDom(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComQIPtr<IHTMLInputElement> spInputElement = spHtmlElement;
	if (spInputElement != NULL)
	{
		// Get the type of the input element.
		CComBSTR bstrInputType;
		HRESULT hRes = spInputElement->get_type(&bstrInputType);
		if ((S_OK == hRes) && (bstrInputType != NULL))
		{
			if (!_wcsicmp(bstrInputType, L"checkbox") || !_wcsicmp(bstrInputType, L"radio"))
			{
				// Radio or checkbox. Search the text at the right.
				return GetRightAccessibleName(spHtmlElement);
			}
			else if (!_wcsicmp(bstrInputType, L"text") || !_wcsicmp(bstrInputType, L"password") || !_wcsicmp(bstrInputType, L"file"))
			{
				// Edit or password. Search the text at the left.
				return GetLeftAccessibleName(spHtmlElement);
			}
		}
		else
		{
			traceLog << "IHTMLInputElement::get_type failed in HtmlHelpers::GetAccessibleNameUsingDom with code:" << hRes << "\n";
			return CComBSTR("");
		}
	}

	// TextArea.
	CComQIPtr<IHTMLTextAreaElement> spTextAreaElement = spHtmlElement;
	if (spTextAreaElement != NULL)
	{
		return GetLeftAccessibleName(spHtmlElement);
	}

	// Select.
	CComQIPtr<IHTMLSelectElement> spSelectElement = spHtmlElement;
	if (spSelectElement != NULL)
	{
		return GetLeftAccessibleName(spHtmlElement);
	}

	return CComBSTR("");
}


CComBSTR HtmlHelpers::GetRightAccessibleName(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);
	CComQIPtr<IHTMLDOMNode> spCurrentNode = spHtmlElement;
	if (spCurrentNode == NULL)
	{
		traceLog << "Can not query for IHTMLDOMNode in HtmlHelpers::GetRightAccessibleName\n";
		return L"";
	}

	while (TRUE)
	{
		// Get the right sibling
		CComQIPtr<IHTMLDOMNode> spRightSiblingNode;
		HRESULT hRes = spCurrentNode->get_nextSibling(&spRightSiblingNode);
		if ((S_OK == hRes) && (spRightSiblingNode != NULL))
		{
			CComBSTR bstrRightText = GetRightText(spRightSiblingNode);
			if (!Common::IsEmptyOrBlank(bstrRightText))
			{
				return bstrRightText;
			}
			else
			{
				spCurrentNode = spRightSiblingNode;
			}
		}
		else
		{
			// The current node has not a rigth sibling. Use the parent node.
			CComQIPtr<IHTMLDOMNode> spParentNode;
			HRESULT hRes = spCurrentNode->get_parentNode(&spParentNode);
			if ((S_OK == hRes) && (spParentNode != NULL))
			{
				spCurrentNode = spParentNode;
			}
			else
			{
				// No more parents.
				break;
			}
		}
	}

	return L"";
}


CComBSTR HtmlHelpers::GetLeftAccessibleName(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);
	CComQIPtr<IHTMLDOMNode> spCurrentNode = spHtmlElement;
	if (spCurrentNode == NULL)
	{
		traceLog << "Can not query for IHTMLDOMNode in HtmlHelpers::GetLeftAccessibleName\n";
		return L"";
	}

	while (TRUE)
	{
		// Get the right sibling
		CComQIPtr<IHTMLDOMNode> spLeftSiblingNode;
		HRESULT hRes = spCurrentNode->get_previousSibling(&spLeftSiblingNode);
		if ((S_OK == hRes) && (spLeftSiblingNode != NULL))
		{
			CComBSTR bstrLeftText = GetRightText(spLeftSiblingNode);
			if (!Common::IsEmptyOrBlank(bstrLeftText))
			{
				return bstrLeftText;
			}
			else
			{
				spCurrentNode = spLeftSiblingNode;
			}
		}
		else
		{
			// The current node has not a left sibling. Use the parent node.
			CComQIPtr<IHTMLDOMNode> spParentNode;
			HRESULT hRes = spCurrentNode->get_parentNode(&spParentNode);
			if ((S_OK == hRes) && (spParentNode != NULL))
			{
				spCurrentNode = spParentNode;
			}
			else
			{
				// No more parents.
				break;
			}
		}
	}

	return L"";
}


CComBSTR HtmlHelpers::GetRightText(CComQIPtr<IHTMLDOMNode> spNode)
{
	ATLASSERT(spNode != NULL);
	const int NODE_TEXT = 3; // According to MSDN.

	// Get the node type.
	long nNodeType;
	HRESULT hRes = spNode->get_nodeType(&nNodeType);
	if ((S_OK == hRes) && (NODE_TEXT == nNodeType))
	{
		// Get the text of the node.
		CComVariant vText;
		hRes = spNode->get_nodeValue(&vText);
		if ((S_OK == hRes) && (VT_BSTR == vText.vt) && (vText.bstrVal != NULL))
		{
			if (!Common::IsEmptyOrBlank(vText.bstrVal))
			{
				return vText.bstrVal;
			}
		}
	}

	// Can not get the text of spNode. Search starting with the first child.
	CComQIPtr<IHTMLDOMNode> spChildNode;
	hRes = spNode->get_firstChild(&spChildNode);
	if ((hRes != S_OK) || (spChildNode == NULL))
	{
		return L"";
	}

	while (spChildNode != NULL)
	{
		CComBSTR bstrText = GetRightText(spChildNode);
		if (!Common::IsEmptyOrBlank(bstrText))
		{
			return bstrText;
		}

		CComQIPtr<IHTMLDOMNode> spNextChild;
		hRes = spChildNode->get_nextSibling(&spNextChild);
		spChildNode = spNextChild;
	}

	return L"";
}


CComBSTR HtmlHelpers::GetLeftText(CComQIPtr<IHTMLDOMNode> spNode)
{
	ATLASSERT(spNode != NULL);
	const int NODE_TEXT = 3; // According to MSDN.

	// Get the node type.
	long nNodeType;
	HRESULT hRes = spNode->get_nodeType(&nNodeType);
	if ((S_OK == hRes) && (NODE_TEXT == nNodeType))
	{
		// Get the text of the node.
		CComVariant vText;
		hRes = spNode->get_nodeValue(&vText);
		if ((S_OK == hRes) && (VT_BSTR == vText.vt) && (vText.bstrVal != NULL))
		{
			if (!Common::IsEmptyOrBlank(vText.bstrVal))
			{
				return vText.bstrVal;
			}
		}
	}

	// Can not get the text of spNode. Search starting with the last child.
	CComQIPtr<IHTMLDOMNode> spChildNode;
	hRes = spNode->get_lastChild(&spChildNode);
	if ((hRes != S_OK) || (spChildNode == NULL))
	{
		return L"";
	}

	while (spChildNode != NULL)
	{
		CComBSTR bstrText = GetLeftText(spChildNode);
		if (!Common::IsEmptyOrBlank(bstrText))
		{
			return bstrText;
		}

		CComQIPtr<IHTMLDOMNode> spPreviousChild;
		hRes = spChildNode->get_previousSibling(&spPreviousChild);
		spChildNode = spPreviousChild;
	}

	return L"";
}


CComQIPtr<IHTMLWindow2> HtmlHelpers::AccessibleToHtmlWindow(CComQIPtr<IAccessible> spAccessible)
{
	ATLASSERT(spAccessible != NULL);

	CComQIPtr<IServiceProvider>	spServProvider = spAccessible;
	if (spServProvider != NULL)
	{
		CComQIPtr<IHTMLWindow2> spHtmlWindow;
		HRESULT hRes = spServProvider->QueryService(IID_IHTMLWindow2, IID_IHTMLWindow2, (void**)&spHtmlWindow);
		if (spHtmlWindow != NULL)
		{
			return spHtmlWindow;
		}
		else
		{
			traceLog << "QueryService failed in HtmlHelpers::AccessibleToHtmlWindow with code:" << hRes << "\n";
		}
	}
	else
	{
		traceLog << "Can not get IServiceProvider from IAccessible in HtmlHelpers::AccessibleToHtmlWindow\n";
	}

	return CComQIPtr<IHTMLWindow2>();
}


CComQIPtr<IHTMLDocument2> HtmlHelpers::AccessibleToHtmlDocument(CComQIPtr<IAccessible> spAccessible)
{
	ATLASSERT(spAccessible != NULL);

	CComQIPtr<IServiceProvider>	spServProvider = spAccessible;
	if (spServProvider != NULL)
	{
		CComQIPtr<IHTMLDocument2> spHtmlDoc;
		HRESULT hRes = spServProvider->QueryService(IID_IHTMLDocument2, IID_IHTMLDocument2, (void**)&spHtmlDoc);
		if (spHtmlDoc != NULL)
		{
			return spHtmlDoc;
		}
		else
		{
			traceLog << "QueryService failed in HtmlHelpers::AccessibleToHtmlDocument with code:" << hRes << "\n";
		}
	}
	else
	{
		traceLog << "Can not get IServiceProvider from IAccessible in HtmlHelpers::AccessibleToHtmlDocument\n";
	}

	return CComQIPtr<IHTMLDocument2>();
}


CComQIPtr<IHTMLElement> HtmlHelpers::AccessibleToHtmlElement(CComQIPtr<IAccessible> spAccessible)
{
	ATLASSERT(spAccessible != NULL);

	CComQIPtr<IServiceProvider>	spServProvider = spAccessible;
	if (spServProvider != NULL)
	{
		CComQIPtr<IHTMLElement>	spHtmlElement;
		HRESULT hRes = spServProvider->QueryService(IID_IHTMLElement, IID_IHTMLElement, (void**)&spHtmlElement);
		if (spHtmlElement != NULL)
		{
			return spHtmlElement;
		}
		else
		{
			traceLog << "QueryService failed in HtmlHelpers::AccessibleToHtmlElement with code:" << hRes << "\n";
		}
	}
	else
	{
		traceLog << "Can not get IServiceProvider from IAccessible in HtmlHelpers::AccessibleToHtmlElement\n";
	}

	return CComQIPtr<IHTMLElement>();
}


CComQIPtr<IAccessible> HtmlHelpers::HtmlElementToAccessible(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComQIPtr<IServiceProvider>	spServProvider = spHtmlElement;
	if (spServProvider != NULL)
	{
		CComQIPtr<IAccessible>	spAccessible;
		HRESULT hRes = spServProvider->QueryService(IID_IAccessible, IID_IAccessible, (void**)&spAccessible);
		if (spAccessible != NULL)
		{
			return spAccessible;
		}
		else
		{
			traceLog << "QueryService failed in HtmlHelpers::HtmlElementToAccessible with code:" << hRes << "\n";
		}
	}
	else
	{
		traceLog << "Can not get IServiceProvider from IHTMLElement in HtmlHelpers::HtmlElementToAccessible\n";
	}

	return CComQIPtr<IAccessible>();
}


CComQIPtr<IAccessible> HtmlHelpers::HtmlDocumentToAccessible(CComQIPtr<IHTMLDocument2> spDocument)
{
	ATLASSERT(spDocument != NULL);

	CComQIPtr<IServiceProvider> spServProvider = spDocument;
	if (spServProvider != NULL)
	{
		CComQIPtr<IAccessible> spAccDocument;
		HRESULT hRes = spServProvider->QueryService(IID_IAccessible, IID_IAccessible, (void**)&spAccDocument);
		if (spAccDocument != NULL)
		{
			return spAccDocument;
		}
		else
		{
			traceLog << "QueryService failed in HtmlHelpers::HtmlDocumentToAccessible with code:" << hRes << "\n";
		}
	}
	else
	{
		traceLog << "Can not get IServiceProvider from IHTMLDocument2 in HtmlHelpers::HtmlDocumentToAccessible\n";
	}

	return CComQIPtr<IAccessible>();
}


// Returns FALSE if any error occurs.
BOOL HtmlHelpers::GetElemClickScreenPoint
	(
		CComQIPtr<IHTMLElement> spElement,
		POINT*                  pClickPoint,
		BOOL                    bGetInputFileButton,
		int                     nZoomLevel
	)
{
	ATLASSERT(spElement   != NULL);
	ATLASSERT(pClickPoint != NULL);

	RECT screenRect = { 0 };
	BOOL  bRes = GetElementScreenLocation(spElement, &screenRect, nZoomLevel);

	if (!bRes)
	{
		traceLog << "GetElementScreenLocation failed in HtmlHelpers::GetElemClickScreenPoint with code\n";
		return FALSE;
	}

	pClickPoint->x = screenRect.left + 1;
	pClickPoint->y = screenRect.top  + 1;

	if (bGetInputFileButton && IsInputFileElement(spElement))
	{
		pClickPoint->x += ((screenRect.right - screenRect.left) * 5 / 6);
	}

	return TRUE;
}


CComQIPtr<IHTMLMapElement> HtmlHelpers::GetAreaMap(CComQIPtr<IHTMLAreaElement> spArea)
{
	ATLASSERT(spArea != NULL);

	CComQIPtr<IHTMLElement> spElement = spArea;
	if (spElement == NULL)
	{
		traceLog << "Query for IHTMLElement failed in HtmlHelpers::GetAreaMap\n";
		return CComQIPtr<IHTMLMapElement>();
	}

	while (TRUE)
	{
		CComQIPtr<IHTMLElement> spParent;
		HRESULT hRes = spElement->get_parentElement(&spParent);
		if (FAILED(hRes))
		{
			traceLog << "IHTMLElement::get_parentElement failed in HtmlHelpers::GetAreaMap\n";
			return CComQIPtr<IHTMLMapElement>();
		}

		if (spParent == NULL)
		{
			// There is no parent <map>.
			return CComQIPtr<IHTMLMapElement>();
		}

		CComQIPtr<IHTMLMapElement> spParentMap = spParent;
		if (spParentMap != NULL)
		{
			// The parent <map> was found.
			return spParentMap;
		}

		spElement = spParent;
	}
}


CComQIPtr<IHTMLImgElement> HtmlHelpers::GetMapImage(CComQIPtr<IHTMLMapElement> spMap)
{
	ATLASSERT(spMap != NULL);

	// Get the map name.
	CComBSTR bstrMapName;
	HRESULT  hRes = spMap->get_name(&bstrMapName);
	if (FAILED(hRes) || (bstrMapName == NULL))
	{
		traceLog << "Can not get the map name in HtmlHelpers::GetMapImage\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	CComQIPtr<IHTMLElement> spElementMap = spMap;
	if (spElementMap == NULL)
	{
		traceLog << "Query for IHTMLElement failed in HtmlHelpers::GetMapImage\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	// Get the document.
	CComQIPtr<IDispatch> spDisp;
	hRes = spElementMap->get_document(&spDisp);

	CComQIPtr<IHTMLDocument2> spDocument = spDisp;
	if (spDocument == NULL)
	{
		traceLog << "Failed to get the document in HtmlHelpers::GetMapImage, code " << hRes << "\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	// Get the images collection in the document.
	CComQIPtr<IHTMLElementCollection> spImages;
	hRes = spDocument->get_images(&spImages);
	if (FAILED(hRes) || (spImages == NULL))
	{
		traceLog << "Can not get the images collection in HtmlHelpers::GetMapImage, code " << hRes << "\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	// Get the number of images in the image collection.
	long nNumberOfImages;
	hRes = spImages->get_length(&nNumberOfImages);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLElementCollection::get_length failed with code " << hRes << " in HtmlHelpers::GetMapImage\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	// Browse the images collection to find the image of the map.
	for (int i = 0; i < nNumberOfImages; ++i)
	{
		// Get the current element in the collection.
		CComQIPtr<IDispatch> spDisp;
		hRes = spImages->item(CComVariant(i), CComVariant(), &spDisp);
		
		CComQIPtr<IHTMLImgElement> spCurrentImage = spDisp;
		if (spCurrentImage == NULL)
		{
			// Can not get the current image in the collection. Skip to next image.
			traceLog << "Failed to get the " << i << "th image element in HtmlHelpers::GetMapImage, code " << hRes << "\n";
			continue;
		}

		// Get the useMap property.
		CComBSTR bstrUseMap;
		hRes = spCurrentImage->get_useMap(&bstrUseMap);
		if (FAILED(hRes) || (bstrUseMap.Length() == 0))
		{
			// Can not get the useMap of the current image. Skip to next image.
			continue;
		}

		if (!wcscmp(bstrMapName, (bstrUseMap + 1)))
		{
			// The image was found !
			return spCurrentImage;
		}
	}

	// The image was not found.
	return CComQIPtr<IHTMLImgElement>();
}


CComQIPtr<IHTMLImgElement> HtmlHelpers::GetAreaImage(CComQIPtr<IHTMLAreaElement> spArea)
{
	ATLASSERT(spArea != NULL);
	CComQIPtr<IHTMLMapElement> spMap = GetAreaMap(spArea);
	if (spMap == NULL)
	{
		traceLog << "Can not get the area's map in HtmlHelpers::GetAreaImage\n";
		return CComQIPtr<IHTMLImgElement>();
	}

	return GetMapImage(spMap);
}

BOOL HtmlHelpers::GetAreaOffset(CComQIPtr<IHTMLAreaElement> spArea, POINT* pOffset)
{
	ATLASSERT(spArea  != NULL);
	ATLASSERT(pOffset != NULL);

	CComQIPtr<IHTMLElement2> spElement2 = spArea;
	if (spElement2 == NULL)
	{
		traceLog << "Query for IHTMLElement2 failed in HtmlHelpers::GetAreaOffset\n";
		return FALSE;
	}

	CComQIPtr<IHTMLRect> spRect;
	HRESULT hRes = spElement2->getBoundingClientRect(&spRect);
	if (FAILED(hRes) || (spRect == NULL))
	{
		traceLog << "IHTMLElement2""getBoundingClientRect failed in HtmlHelpers::GetAreaOffset with code " << hRes << "\n";
		return FALSE;
	}

	long nLeft;
	hRes = spRect->get_left(&nLeft);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLRect::get_left failed in HtmlHelpers::GetAreaOffset with code " << hRes << "\n";
		return FALSE;
	}

	long nTop;
	hRes = spRect->get_top(&nTop);
	if (FAILED(hRes))
	{
		traceLog << "IHTMLRect::get_top failed in HtmlHelpers::GetAreaOffset with code " << hRes << "\n";
		return FALSE;
	}

	// Take care that there are non-rectangular <area>s.
	pOffset->x = nLeft + 1;
	pOffset->y = nTop  + 1;

	return TRUE;
}


BOOL HtmlHelpers::IsInputFileElement(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComQIPtr<IHTMLInputElement> spInputElement = spHtmlElement;
	if (spInputElement == NULL)
	{
		return FALSE;
	}

	CComBSTR bstrInputType;
	HRESULT  hRes = spInputElement->get_type(&bstrInputType);
	if (FAILED(hRes) || (bstrInputType == NULL))
	{
		return FALSE;
	}

	return !_wcsicmp(bstrInputType, L"file");
}


CComQIPtr<IHTMLElement> HtmlHelpers::GetHtmlElementFromScreenPos
							(CComQIPtr<IHTMLDocument2> spDocument, long lX, long lY)
{
	ATLASSERT(spDocument != NULL);

	CComQIPtr<IAccessible> spAccDocument = HtmlHelpers::HtmlDocumentToAccessible(spDocument);
	if (spAccDocument == NULL)
	{
		traceLog << "Can not get IAccessible from IHTMLDocument2 in GetHtmlElementFromScreenPos\n";
		return CComQIPtr<IHTMLElement>();
	}

	long nScreenDocLeft = 0, nScreenDocTop = 0, nScreenDocWidth = 0, nScreenDocHeight = 0;
	HRESULT hRes = spAccDocument->accLocation(&nScreenDocLeft,  &nScreenDocTop, 
	                                          &nScreenDocWidth, &nScreenDocHeight, CComVariant(CHILDID_SELF));
	if (hRes != S_OK)
	{
		traceLog << "IAccessible::accLocation failed with error code: " << hRes << " in GetHtmlElementFromScreenPos\n";
		return CComQIPtr<IHTMLElement>();
	}

	// Get html element from document coordinates.
	CComQIPtr<IHTMLElement> spElementHit;
	hRes = spDocument->elementFromPoint((lX - nScreenDocLeft), (lY - nScreenDocTop), &spElementHit);
	if (spElementHit == NULL)
	{
		traceLog << "IHTMLElement::elementFromPoint failed with code " << hRes << " in HtmlHelpers::GetHtmlElementFromScreenPos\n";
		return CComQIPtr<IHTMLElement>();
	}

	return spElementHit;
}


BOOL HtmlHelpers::GetElementScreenLocation(IHTMLElement* pElement, RECT* pRect, int nZoomLevel)
{
	ATLASSERT(pElement != NULL);
	ATLASSERT(pRect    != NULL);

	// Query for IHTMLElement2 interface.
	CComQIPtr<IHTMLElement2> spElement2 = pElement;
	if (spElement2 == NULL)
	{
		traceLog << "Query for IHTMLElement2 failed in HtmlHelpers::GetElementScreenLocation\n";
		return FALSE;
	}

	// Get coordinates of the element relative to the container document.
	CComQIPtr<IHTMLRect> spHtmlRect;
	HRESULT hRes = spElement2->getBoundingClientRect(&spHtmlRect);
	if ((FAILED(hRes)) || (spHtmlRect == NULL))
	{
		traceLog << "IHTMLElement2::getBoundingClientRect failed with code " << hRes << " in GetElementScreenLocation\n";
		return FALSE;
	}

	if (FAILED(spHtmlRect->get_left(&(pRect->left)))   || 
	    FAILED(spHtmlRect->get_right(&(pRect->right))) ||
	    FAILED(spHtmlRect->get_top(&(pRect->top)))     ||
	    FAILED(spHtmlRect->get_bottom(&(pRect->bottom))))
	{
		traceLog << "Can not get the IHTMLRect object in GetElementScreenLocation\n";
		return FALSE;
	}

	// Get the parent document of the html element.
	CComQIPtr<IDispatch> spDisp;
	hRes = pElement->get_document(&spDisp);

	CComQIPtr<IHTMLDocument2> spDocument = spDisp;
	if (spDocument == NULL)
	{
		traceLog << "Can not get IHTMLDocument2 in GetElementScreenLocation\n";
		return FALSE;
	}

	CComQIPtr<IAccessible> spAccDocument = HtmlHelpers::HtmlDocumentToAccessible(spDocument);
	if (spAccDocument == NULL)
	{
		traceLog << "Can not get IAccessible from IHTMLDocument2 in GetElementScreenLocation\n";
		return FALSE;
	}

	long nScreenDocLeft = 0, nScreenDocTop = 0, nScreenDocWidth = 0, nScreenDocHeight = 0;
	hRes = spAccDocument->accLocation(&nScreenDocLeft, &nScreenDocTop, 
	                                          &nScreenDocWidth, &nScreenDocHeight, CComVariant(CHILDID_SELF));
	if (hRes != S_OK)
	{
		traceLog << "IAccessible::accLocation failed with error code: " << hRes << " in GetElementScreenLocation\n";
		return false;
	}

	pRect->left   = (LONG)((pRect->left   * nZoomLevel) / 100.0 + nScreenDocLeft);
	pRect->right  = (LONG)((pRect->right  * nZoomLevel) / 100.0 + nScreenDocLeft);
	pRect->top    = (LONG)((pRect->top    * nZoomLevel) / 100.0 + nScreenDocTop);
	pRect->bottom = (LONG)((pRect->bottom * nZoomLevel) / 100.0 + nScreenDocTop);

	return TRUE;
}


CComBSTR HtmlHelpers::GetTextAttributeValue(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComBSTR bstrResultText;
	if (IsValidForUsingInnerText(spHtmlElement))
	{
		// First try to get the inner text of the HTML element.
		HRESULT  hRes = spHtmlElement->get_innerText(&bstrResultText);
		if ((hRes == S_OK) && (bstrResultText.Length() > 0))
		{
			return bstrResultText;
		}
		else
		{
			// <select> options can have emtpy inner text but a label attribute.
			CComQIPtr<IHTMLOptionElement3> spOption3 = spHtmlElement;
			if (spOption3 != NULL)
			{
				CComBSTR bstrLabelText;
				hRes = spOption3->get_label(&bstrLabelText);

				if ((hRes == S_OK) && (bstrLabelText.Length() > 0))
				{
					return bstrLabelText;
				}
			}
		}
	}

	if (IsValidForComputingTextUsingAA(spHtmlElement))
	{
		// Try to get the text using Accessible technology.
		CComQIPtr<IAccessible> spAccessible = HtmlElementToAccessible(spHtmlElement);
		if (spAccessible != NULL)
		{
			HRESULT hRes = spAccessible->get_accName(CComVariant(CHILDID_SELF), &bstrResultText);
			if ((S_OK == hRes) && (bstrResultText.Length() > 0))
			{
				return bstrResultText;
			}
		}
	}

	// Active Accessibility doesn't provide a name. Get the name using HTML DOM.
	return GetAccessibleNameUsingDom(spHtmlElement);
}


BOOL HtmlHelpers::IsValidForUsingInnerText(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComQIPtr<IHTMLSelectElement> spSelectElement = spHtmlElement;
	if (spSelectElement != NULL)
	{
		return FALSE;
	}

	CComQIPtr<IHTMLTextAreaElement> spTextAreaElement = spHtmlElement;
	if (spTextAreaElement != NULL)
	{
		return FALSE;
	}

	return TRUE;
}


BOOL HtmlHelpers::IsValidForComputingTextUsingAA(CComQIPtr<IHTMLElement> spHtmlElement)
{
	ATLASSERT(spHtmlElement != NULL);

	CComQIPtr<IHTMLInputElement> spInputElement = spHtmlElement;
	if (spInputElement != NULL)
	{
		// Get the type of the input element.
		CComBSTR bstrInputType;
		HRESULT hRes = spInputElement->get_type(&bstrInputType);
		if (bstrInputType == NULL)
		{
			bstrInputType = L"";
		}

		return (!_wcsicmp(bstrInputType, L"checkbox") || !_wcsicmp(bstrInputType, L"radio")    ||
		        !_wcsicmp(bstrInputType, L"text")     || !_wcsicmp(bstrInputType, L"password") ||
				!_wcsicmp(bstrInputType, L"button")   || !_wcsicmp(bstrInputType, L"submit")   ||
				!_wcsicmp(bstrInputType, L"reset")    || !_wcsicmp(bstrInputType, L"file"));
	}

	CComQIPtr<IHTMLTextAreaElement> spTextAreaElement = spHtmlElement;
	if (spTextAreaElement != NULL)
	{
		return TRUE;
	}

	CComQIPtr<IHTMLSelectElement> spSelectElement = spHtmlElement;
	if (spSelectElement != NULL)
	{
		return TRUE;
	}

	CComQIPtr<IHTMLButtonElement> spButton = spHtmlElement;
	if (spButton != NULL)
	{
		return TRUE;
	}

	CComQIPtr<IHTMLImgElement> spImg = spHtmlElement;
	if (spImg != NULL)
	{
		// IMG elements could have title.
		return TRUE;
	}

	CComQIPtr<IHTMLAnchorElement> spAnchor = spHtmlElement;
	if (spAnchor != NULL)
	{
		// Anchors without text could have title.
		return TRUE;
	}

	return FALSE;
}


HRESULT HtmlHelpers::GetIEServerWndFromElement(IHTMLElement* pElement, HWND* pIeServerWnd)
{
	ATLASSERT(pElement     != NULL);
	ATLASSERT(pIeServerWnd != NULL);

	*pIeServerWnd = NULL;

	CComQIPtr<IAccessible> spAcc = HtmlHelpers::HtmlElementToAccessible(pElement);
	if (spAcc == NULL)
	{
		traceLog << "Can not get IAccessible object in HtmlHelpers::GetIEServerWndFromElement\n";
		return E_FAIL;
	}

	HWND    hIeWnd = NULL;
	HRESULT hRes   = ::WindowFromAccessibleObject(spAcc, &hIeWnd);
	if (FAILED(hRes) || !::IsWindow(hIeWnd))
	{
		traceLog << "WindowFromAccessibleObject faied with code " << hRes << " in HtmlHelpers::GetIEServerWndFromElement\n";
		return hRes;
	}

	while (hIeWnd != NULL)
	{
		if (Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"))
		{
			*pIeServerWnd = hIeWnd;
			break;
		}

		hIeWnd = ::GetParent(hIeWnd);
	}

	return S_OK;
}


// Returns top level window for IE6, "TabWindowClass" for IE7.
// TODO: test if it works with non top IWebBrowser objects.
HWND HtmlHelpers::GetIEWndFromBrowser(CComQIPtr<IWebBrowser2> spBrowser)
{
	if (spBrowser == NULL)
	{
		traceLog << "Invalid parameter in HtmlHelpers::GetIEWndFromBrowser\n";
		return NULL;
	}

	HWND hwndIETab = NULL;
	CComQIPtr<IServiceProvider> spServiceProvider = spBrowser;

	if (spServiceProvider != NULL)
	{
		CComQIPtr<IOleWindow> spWindow;
		if (SUCCEEDED(spServiceProvider->QueryService(SID_SShellBrowser, IID_IOleWindow, (void**)&spWindow)))
		{
			spWindow->GetWindow(&hwndIETab);
		}
	}

	if (NULL == hwndIETab)
	{
		traceLog << "hwndIETab is NULL in HtmlHelpers::GetIEWndFromBrowser\n";
	}

	return hwndIETab;
}


// Returns false if spSelectElem is a list (single or multiple selection).
// On error throw an exception.
BOOL HtmlHelpers::IsComboControl(CComQIPtr<IHTMLSelectElement> spSelectElem)
{
	ATLASSERT(spSelectElem != NULL);

	long	nRowsNumber = 0;
	HRESULT hRes = spSelectElem->get_size(&nRowsNumber);
	
	if (FAILED(hRes))
	{
		throw CreateException(_T("IHTMLSelectElement::get_size failed in HtmlHelpers::IsComboControl"));
	}

	return nRowsNumber <= 1; // For combos size is one. On error assume is combo.
}


// Returns true if spSelectElem is a multiple selection list.
// On error throw an exception.
BOOL HtmlHelpers::IsMultipleListControl(CComQIPtr<IHTMLSelectElement> spSelectElem)
{
	ATLASSERT(spSelectElem != NULL);
	ATLASSERT(!IsComboControl(spSelectElem));

	VARIANT_BOOL vbMultiple = VARIANT_FALSE;
	HRESULT      hRes = spSelectElem->get_multiple(&vbMultiple);
	
	if (FAILED(hRes))
	{
		throw CreateException(_T("IHTMLSelectElement::get_multiple failed in HtmlHelpers::IsMultipleListControl"));
	}

	return (VARIANT_TRUE == vbMultiple);
}


// Returns FALSE on failure.
BOOL HtmlHelpers::FireEventOnElement
	(
		CComQIPtr<IHTMLElement> spTargetElement,
		CComBSTR                bstrEvName,
		WCHAR                   wchCharToRise,
		LONG                    x,
		LONG                    y,
		BOOL                    bRightClick)
{
	ATLASSERT(spTargetElement    != NULL);
	ATLASSERT(bstrEvName.Length() > 0);

	// Get parent document of spTargetElement.
	CComQIPtr<IHTMLDocument2> spParentDoc;
	CComQIPtr<IDispatch>      spDispDoc;
	HRESULT                   hRes = spTargetElement->get_document(&spDispDoc);
	
	spParentDoc = spDispDoc;
	if (spParentDoc == NULL)
	{
		traceLog << "Can not get spParentDoc in HtmlHelpers::FireEventOnElement, code: " << hRes << "\n";
		return FALSE;
	}

	// IHTMLDocument4 interface needed.
	CComQIPtr<IHTMLDocument4> spDoc4 = spParentDoc;
	if (spDoc4 == NULL)
	{
		traceLog << "IHTMLDocument4 not supported in HtmlHelpers::FireEventOnElement\n";
		return FALSE;
	}

	// Create a new event based on the provided "bstrEvName" event name.
	CComQIPtr<IHTMLEventObj> spEvent;
	hRes = spDoc4->createEventObject(NULL, &spEvent);
	if (FAILED(hRes) || (spEvent == NULL))
	{
		traceLog << "Can not create event in HtmlHelpers::FireEventOnElement, code: " << hRes << "\n";
		return FALSE;
	}

	if (bRightClick)
	{
		CComQIPtr<IHTMLEventObj2> spEvent2 = spEvent;

		if (spEvent2 == NULL)
		{
			traceLog << "IHTMLEventObj2 not supported in HtmlHelpers::FireEventOnElement\n";
			return FALSE;
		}

		hRes = spEvent2->put_button(2); // Right button is pressed.
		if (FAILED(hRes))
		{
			traceLog << "put_button failed in HtmlHelpers::FireEventOnElement\n";
			return FALSE;
		}
	}

	// IHTMLElement3 interface needed.
	CComQIPtr<IHTMLElement3> spElem3 = spTargetElement;
	if (spElem3 == NULL)
	{
		traceLog << "IHTMLElement3 not supported in HtmlHelpers::FireEventOnElement\n";
		return FALSE;
	}

	if (wchCharToRise != L'\0')
	{
		hRes = spEvent->put_keyCode(wchCharToRise);
		if (FAILED(hRes))
		{
			traceLog << "IHTMLEventObj::put_keyCode failed in HtmlHelpers::FireEventOnElement with code " << hRes << "\n";
			return FALSE;
		}
	}

	if (-1 == x)
	{
		x = 1;
	}

	if (-1 == y)
	{
		y = 1;
	}

	if ((x != -1) && (y != -1))
	{
		// TODO: set x, y, clientX, clientY, screenX, screenY

		CComQIPtr<IHTMLEventObj2> spEvent2 = spEvent;
		if (spEvent2 == NULL)
		{
			traceLog << "Can not get IHTMLEventObj2 in HtmlHelpers::FireEventOnElement with code " << hRes << "\n";
			return FALSE;
		}

		hRes = spEvent2->put_offsetX(x);
		if (FAILED(hRes))
		{
			traceLog << "IHTMLEventObj::put_offsetX failed in HtmlHelpers::FireEventOnElement with code " << hRes << "\n";
			return FALSE;
		}

		hRes = spEvent2->put_offsetY(y);
		if (FAILED(hRes))
		{
			traceLog << "IHTMLEventObj::put_offsetY failed in HtmlHelpers::FireEventOnElement with code " << hRes << "\n";
			return FALSE;
		}
	}

	// Fire event on spElem3.
	VARIANT_BOOL vbCancel;
	CComVariant  varEvnt(spEvent);
	hRes = spElem3->fireEvent(bstrEvName, &varEvnt, &vbCancel);

	if (SUCCEEDED(hRes))
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


CComQIPtr<IHTMLElement> HtmlHelpers::GetAncestorByTag(IHTMLElement* pElement, CComBSTR bstrTagNameToFind)
{
	ATLASSERT(pElement          != NULL);
	ATLASSERT(bstrTagNameToFind != NULL);

	CComQIPtr<IHTMLElement>	 spCrntElem = pElement;
	while (spCrntElem != NULL)
	{
		// Get the tag name of the current element.
		CComBSTR bstrTagName;
		HRESULT hRes = spCrntElem->get_tagName(&bstrTagName);
		if (FAILED(hRes) || (0 == bstrTagName.Length()))
		{
			break;
		}

		if (!_wcsicmp(bstrTagNameToFind, bstrTagName))
		{
			return spCrntElem;
		}

		CComQIPtr<IHTMLElement> spParent;
		hRes = spCrntElem->get_parentElement(&spParent);
		spCrntElem = spParent;

		if (FAILED(hRes))
		{
			break;
		}
	}

	return CComQIPtr<IHTMLElement>();
}


BOOL HtmlHelpers::IsWindowReady(CComQIPtr<IHTMLWindow2> spWindow)
{
	ATLASSERT(spWindow != NULL);

	CComQIPtr<IHTMLDocument2> spDoc = HtmlHelpers::HtmlWindowToHtmlDocument(spWindow);
	if (spDoc == NULL)
	{
		throw CreateException(_T("HtmlHelpers::HtmlWindowToHtmlDocument failed in HtmlHelpers::IsWindowReady"));
	}

	CComBSTR bstrState;
	HRESULT  hRes = spDoc->get_readyState(&bstrState);

	if (FAILED(hRes) || (bstrState == NULL))
	{
		throw CreateException(_T("IHTMLDocument2::get_readyState failed in HtmlHelpers::IsWindowReady"));
	}

	// On digg.com a document remains forever in interactive state.
	if (_wcsicmp(bstrState, L"complete") && _wcsicmp(bstrState, L"interactive"))
	{
		// Document is not complete yet.
		return FALSE;
	}

	// Get the subframes collection of the spWindow html window.
	CComQIPtr<IHTMLFramesCollection2> spFrameCollection;
	hRes = spWindow->get_frames(&spFrameCollection);

	if (FAILED(hRes) || (spFrameCollection == NULL))
	{
		throw CreateException(_T("IHTMLWindow2::get_frames failed in HtmlHelpers::IsWindowReady"));
	}

	// Get the number of frames in the spFrameCollection.
	long nNumberOfSubframes = 0;
	hRes = spFrameCollection->get_length(&nNumberOfSubframes);

	if (FAILED(hRes))
	{
		throw CreateException(_T("IHTMLFramesCollection2::get_length failed in HtmlHelpers::IsWindowReady"));
	}

	// Browse the subframes collection.
	for (long i = 0; i < nNumberOfSubframes; ++i)
	{
		CComVariant		varSubframe;
		CComVariant		varIndex(i);
		hRes = spFrameCollection->item(&varIndex, &varSubframe);
		if (hRes!= S_OK)
		{
			throw CreateException(_T("IHTMLFramesCollection2::item failed in HtmlHelpers::IsWindowReady"));
		}

		if ((varSubframe.vt != VT_DISPATCH) || (NULL == varSubframe.pdispVal))
		{
			throw CreateException(_T("IHTMLFramesCollection2::item returned invalid variant in HtmlHelpers::IsWindowReady"));
		}

		// Get the current subframe;
		CComQIPtr<IHTMLWindow2>	spCurrentFrame = varSubframe.pdispVal;
		if (spCurrentFrame == NULL)
		{
			throw CreateException(_T("Query for IHTMLWindow2 failed in HtmlHelpers::IsWindowReady"));
		}

		if (!IsWindowReady(spCurrentFrame))
		{
			return FALSE;
		}
	}

	return TRUE;
}


BOOL HtmlHelpers::IsBrowserReady(CComQIPtr<IWebBrowser2> spBrws)
{
	ATLASSERT(spBrws != NULL);

	// Get the browswer ready state.
	READYSTATE state;
	HRESULT    hRes = spBrws->get_ReadyState(&state);
	if (FAILED(hRes))
	{
		throw CreateException(_T("IWebBrowser2::get_ReadyState failed in HtmlHelpers::IsBrowserReady"));
	}

	if (READYSTATE_COMPLETE != state)
	{
		// The browser is still loading.
		return FALSE;
	}

	CComQIPtr<IHTMLWindow2> spTopWindow = HtmlHelpers::HtmlWebBrowserToHtmlWindow(spBrws);
	if (spTopWindow == NULL)
	{
		throw CreateException(_T("HtmlHelpers::HtmlWebBrowserToHtmlWindow failed in HtmlHelpers::IsBrowserReady"));
	}

	// Is this a good idea?? Problems on digg.com where a document in a frame remained forever in interactive state.
	// Recursivelly check each document in hierarchy.
	return IsWindowReady(spTopWindow);
}


HRESULT HtmlHelpers::RightClickElementAt(IHTMLElement* pElement, LONG x, LONG y)
{
	if (NULL == pElement)
	{
		return E_INVALIDARG;
	}

	const TCHAR* clickEvents[] = 
	{
		_T("onmousemove"), _T("onmouseover"), _T("onmouseenter"), _T("onmousedown"), _T("onbeforeactivate"),
		_T("onactivate"),  _T("onfocusin"),   _T("onfocus"),      _T("onmouseup"),   _T("oncontextmenu"),
	};

	const int numberOfEvents = sizeof(clickEvents) / sizeof(clickEvents[0]);

	BOOL bRes = TRUE;
	for (int i = 0; i < numberOfEvents; ++i)
	{
		BOOL bEventRised = HtmlHelpers::FireEventOnElement(pElement, clickEvents[i], L'\0', x, y, TRUE);
		if (!bEventRised)
		{
			bRes = FALSE;
			traceLog << "Failed to rise " << clickEvents[i] << " event in HtmlHelpers::RightClickElementAt\n";
		}
	}

	return bRes ? S_OK : E_FAIL;
}


HRESULT HtmlHelpers::ClickElementAt(IHTMLElement* pElement, LONG x, LONG y, BOOL bRightClick)
{
	if (NULL == pElement)
	{
		return E_INVALIDARG;
	}

	if (!bRightClick)
	{
		const TCHAR* clickEvents[] = 
		{
			_T("onmousemove"), _T("onmouseover"), _T("onmouseenter"), _T("onmousedown"), _T("onbeforeactivate"),
			_T("onactivate"),  _T("onfocusin"),   _T("onfocus"),      _T("onmouseup"),
		};

		const int numberOfEvents = sizeof(clickEvents) / sizeof(clickEvents[0]);

		BOOL bRes = TRUE;
		for (int i = 0; i < numberOfEvents; ++i)
		{
			BOOL bEventRised = HtmlHelpers::FireEventOnElement(pElement, clickEvents[i], L'\0', x, y);
			if (!bEventRised)
			{
				bRes = FALSE;
				traceLog << "Failed to rise " << clickEvents[i] << " event in HtmlHelpers::ClickElementAt\n";
			}
		}

		if (NeedFireOnclick(pElement))
		{
			BOOL bEventRised = HtmlHelpers::FireEventOnElement(pElement, L"onclick", L'\0', x, y);
			if (!bEventRised)
			{
				bRes = FALSE;
				traceLog << "Failed to rise onclick event in HtmlHelpers::ClickElementAt\n";
			}
		}
		else
		{
			HRESULT hRes = pElement->click();
			return hRes;
		}


		return bRes ? S_OK : E_FAIL;
	}
	else
	{
		return HtmlHelpers::RightClickElementAt(pElement, x, y);
	}
}


BOOL HtmlHelpers::NeedFireOnclick(IHTMLElement* pElement)
{
	ATLASSERT(pElement != NULL);

	CComBSTR bstrTagName;
	HRESULT  hRes = pElement->get_tagName(&bstrTagName);

	if (FAILED(hRes))
	{
		traceLog << "IHTMLElement::get_tagName failed in HtmlHelpers::FireOnclick with code " << hRes <<  "\n";
		return FALSE;
	}

	if (!_wcsicmp(bstrTagName, L"INPUT") || !_wcsicmp(bstrTagName, L"AREA")   ||
	    !_wcsicmp(bstrTagName, L"A")     || !_wcsicmp(bstrTagName, L"SUBMIT") ||
		!_wcsicmp(bstrTagName, L"LABEL") || !_wcsicmp(bstrTagName, L"BUTTON"))
	{
		return FALSE;
	}

	CComQIPtr<IHTMLElement> spAnchorAncestor = HtmlHelpers::GetAncestorByTag(pElement, L"A");
	return (spAnchorAncestor == NULL);
}


CComQIPtr<IHTMLElement> HtmlHelpers::GetFrameElement(CComQIPtr<IHTMLWindow2> spFrame)
{
	CComQIPtr<IHTMLWindow4> spWindow = spFrame;
	if (spWindow == NULL)
	{
		traceLog << "Can NOT get IHTMLWindow4 in HtmlHelpers::GetFrameElement\n";
		return CComQIPtr<IHTMLElement>();
	}

	CComQIPtr<IHTMLFrameBase> spFrameBase;
	HRESULT hRes = spWindow->get_frameElement(&spFrameBase);
	if (SUCCEEDED(hRes))
	{
		CComQIPtr<IHTMLElement> spFrameElement = spFrameBase;
		return spFrameElement;
	}
	else
	{
		CComQIPtr<IHTMLWindow2> spParentWindow;

		spFrame->get_parent(&spParentWindow);
		if (spParentWindow == NULL)
		{
			traceLog << "Can NOT get spParentWindow in HtmlHelpers::GetFrameElement\n";
			return CComQIPtr<IHTMLElement>();
		}

		CComQIPtr<IHTMLDocument3> spDoc = HtmlWindowToHtmlDocument(spFrame);
		if (spDoc == NULL)
		{
			traceLog << "Can NOT get IHTMLDocument3 in HtmlHelpers::GetFrameElement\n";
			return CComQIPtr<IHTMLElement>();
		}

		CComQIPtr<IHTMLElement> spRootElem;
		spDoc->get_documentElement(&spRootElem);

		if (spRootElem == NULL)
		{
			traceLog << "Can NOT get root element in HtmlHelpers::GetFrameElement\n";
			return CComQIPtr<IHTMLElement>();
		}

		// Get current time.
		DWORD                 dwCrntTime = ::GetTickCount();
		Common::Ostringstream outputStream;

		outputStream << dwCrntTime;

		CComBSTR bstrCrntTime = outputStream.str().c_str();

		// Set current time as a custom attribute of the root element.
		HRESULT hRes = spRootElem->setAttribute(CComBSTR(SEARCH_FRAME_ELEMENT_ATTR_NAME), CComVariant(bstrCrntTime), 0);
		if (FAILED(hRes))
		{
			traceLog << "Can NOT get setAttribute in HtmlHelpers::GetFrameElement\n";
			return CComQIPtr<IHTMLElement>();
		}

		return FindFrameElementByTickCount(spParentWindow, bstrCrntTime);
	}
}


CComQIPtr<IHTMLElement> HtmlHelpers::FindFrameElementByTickCount(CComQIPtr<IHTMLWindow2> spWindow, const CComBSTR& bstrCrntTime)
{
	ATLASSERT(spWindow != NULL);

	// Get the html document of the html window.
	CComQIPtr<IHTMLDocument2> spDoc = HtmlHelpers::HtmlWindowToHtmlDocument(spWindow);
	if (spDoc == NULL)
	{
		traceLog << "Can not get IHTMLDocument2 in HtmlHelpers::FindFrameElementByTickCount\n";
		return CComQIPtr<IHTMLElement>();
	}

	// Get the "all" collection of the html document.
	CComQIPtr<IHTMLElementCollection> spAllCollection;
	HRESULT hRes = spDoc->get_all(&spAllCollection);
	if ((hRes != S_OK) || (spAllCollection == NULL))
	{
		traceLog << "Can not get the all collection of the document in HtmlHelpers::FindFrameElementByTickCount\n";
		return CComQIPtr<IHTMLElement>();
	}

	CComQIPtr<IHTMLElementCollection> spFrameTagCollection = HtmlHelpers::GetElementCollectionByTag(spAllCollection, CComBSTR("IFRAME"));
	CComQIPtr<IHTMLElement> spResult = FindFrameElementByTickCount(spFrameTagCollection, bstrCrntTime);

	if (spResult == NULL)
	{
		CComQIPtr<IHTMLElementCollection> spIFrameTagCollection = HtmlHelpers::GetElementCollectionByTag(spAllCollection, CComBSTR("FRAME"));
		spResult = FindFrameElementByTickCount(spIFrameTagCollection, bstrCrntTime);
	}

	return spResult;
}


CComQIPtr<IHTMLElementCollection> HtmlHelpers::GetElementCollectionByTag(CComQIPtr<IHTMLElementCollection> spCollection,
                                                                                   const CComBSTR& bstrTagName)
{
	ATLASSERT(spCollection != NULL);
	ATLASSERT(bstrTagName  != NULL);

	CComQIPtr<IDispatch> spDisp;
	HRESULT hRes = spCollection->tags(CComVariant(bstrTagName), &spDisp);
	if ((hRes != S_OK) || (spDisp == NULL))
	{
		traceLog << "IHTMLElementCollection::tags failed in HtmlHelpers::GetElementCollectionByTag\n";
		return CComQIPtr<IHTMLElementCollection>();
	}

	CComQIPtr<IHTMLElementCollection> spResultCollection = spDisp;
	if (spResultCollection == NULL)
	{
		traceLog << "Query for interface IHTMLElementCollection failed in HtmlHelpers::GetElementCollectionByTag\n";
		return CComQIPtr<IHTMLElementCollection>();
	}

	return spResultCollection;
}


CComQIPtr<IHTMLElement> HtmlHelpers::FindFrameElementByTickCount
	(CComQIPtr<IHTMLElementCollection> spElemCollection, const CComBSTR& bstrCrntTime)
{
	if (spElemCollection == NULL)
	{
		traceLog << "spElemCollection is NULL in HtmlHelpers::FindFrameElementByTickCount\n";
		return CComQIPtr<IHTMLElement>();
	}

	// Get the size of the tag collection.
	long    nCollectionSize;
	HRESULT hRes = spElemCollection->get_length(&nCollectionSize);

	if (hRes != S_OK)
	{
		traceLog << "IHTMLElementCollection::get_length failed in HtmlHelpers::FindFrameElementByTickCount\n";
		return CComQIPtr<IHTMLElement>();
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		CComQIPtr<IDispatch> spItemDisp;
		hRes = spElemCollection->item(CComVariant(i), CComVariant(), &spItemDisp);
		if ((hRes != S_OK) || (spItemDisp == NULL))
		{
			traceLog << "Can not get the" << i << "th item in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		CComQIPtr<IHTMLFrameBase2> spCrntFrameBase = spItemDisp;
		if (spCrntFrameBase == NULL)
		{
			traceLog << "Query IHTMLFrameBase2 from IDispatch failed for index" << i << " in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		CComQIPtr<IHTMLWindow2> spCrntWindow;
		spCrntFrameBase->get_contentWindow(&spCrntWindow);

		if (spCrntWindow == NULL)
		{
			traceLog << "get_contentWindow failed in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		CComQIPtr<IHTMLDocument3> spCrntDoc = HtmlWindowToHtmlDocument(spCrntWindow);
		if (spCrntDoc == NULL)
		{
			traceLog << "Can NOT get IHTMLDocument3 in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		CComQIPtr<IHTMLElement> spRootElem;
		spCrntDoc->get_documentElement(&spRootElem);

		if (spRootElem == NULL)
		{
			traceLog << "Can NOT get root element in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		CComVariant vTickAttr;
		HRESULT     hRes = spRootElem->getAttribute(CComBSTR(SEARCH_FRAME_ELEMENT_ATTR_NAME), 1, &vTickAttr);

		if (FAILED(hRes) || (vTickAttr.vt != VT_BSTR) || (NULL == vTickAttr.bstrVal))
		{
			traceLog << "getAttribute failed in HtmlHelpers::FindFrameElementByTickCount\n";
			continue;
		}

		if (bstrCrntTime == vTickAttr.bstrVal)
		{
			CComQIPtr<IHTMLElement> spResult = spCrntFrameBase;
			return spResult;
		}
	}

	return CComQIPtr<IHTMLElement>();
}


CComQIPtr<IWebBrowser2> HtmlHelpers::GetBrowserFromIEServerWnd(HWND hIeWnd)
{
	if (!::IsWindow(hIeWnd) || (Common::GetWndClass(hIeWnd) != _T("Internet Explorer_Server")))
	{
		return CComQIPtr<IWebBrowser2>();
	}

	// Get the accessible object from the "Internet Explorer_Server" window.
	CComQIPtr<IAccessible> spAccObj;
	HRESULT hRes = ::AccessibleObjectFromWindow(hIeWnd, OBJID_WINDOW, IID_IAccessible, (void**)&spAccObj);

	if (FAILED(hRes) || (spAccObj == NULL))
	{
		traceLog << "AccessibleObjectFromWindow failed in HtmlHelpers::GetBrowserFromIEServerWndD code:" << hRes << "\n";
		return CComQIPtr<IWebBrowser2>();
	}

	// Get the service provider from the accessible object.
	CComQIPtr<IServiceProvider>	spServProv = spAccObj;
	if (spServProv == NULL)
	{
		traceLog << "Can NOT get IServiceProvider in HtmlHelpers::GetBrowserFromIEServerWndD\n";
		return CComQIPtr<IWebBrowser2>();
	}

	// Query the service provider for IHTMLWindow2 interface
	CComQIPtr<IHTMLWindow2> spHtmlWindow;
	hRes = spServProv->QueryService(IID_IHTMLWindow2, IID_IHTMLWindow2, (void**)&spHtmlWindow);
	
	if (FAILED(hRes) || (spHtmlWindow == NULL))
	{
		traceLog << "Can NOT get IHTMLWindow2 in HtmlHelpers::GetBrowserFromIEServerWndD\n";
		return CComQIPtr<IWebBrowser2>();
	}

	// Get the top parent window.
	CComQIPtr<IHTMLWindow2>	spTopWindow;
	hRes = spHtmlWindow->get_top(&spTopWindow);

	if (FAILED(hRes) || (spTopWindow == NULL))
	{
		traceLog << "Can NOT get top window in HtmlHelpers::GetBrowserFromIEServerWndD\n";
		return CComQIPtr<IWebBrowser2>();
	}

	return HtmlWindowToHtmlWebBrowser(spTopWindow);
}


HWND HtmlHelpers::GetIEServerFromScrPt(LONG x, LONG y)
{
	// First find the top parent from screen point.
	POINT pt = { x, y };
	HWND  hResult  = NULL;
	HWND  hCrntWnd = ::ChildWindowFromPointEx(::GetDesktopWindow(), pt, CWP_SKIPTRANSPARENT | CWP_SKIPINVISIBLE);

	if (!hCrntWnd)
	{
		traceLog << "Can NOT get top window in HtmlHelpers::GetIEServerFromScrPt\n";
		return hResult;
	}

	::ScreenToClient(hCrntWnd, &pt);

	// Search for "Internet Explorer_Server" child.
	while (TRUE)
	{
		HWND hChildWnd = ::ChildWindowFromPointEx(hCrntWnd, pt, CWP_SKIPTRANSPARENT | CWP_SKIPINVISIBLE);
		if (!hChildWnd)
		{
			traceLog << "Can NOT get child window in HtmlHelpers::GetIEServerFromScrPt\n";
			break;
		}

		if (Common::GetWndClass(hChildWnd) == _T("Internet Explorer_Server"))
		{
			// We have found what we are looking for.
			hResult = hChildWnd;
			break;
		}

		if (hChildWnd == hCrntWnd)
		{
			// The end, no more children.
			break;
		}

		BOOL bRes = ::MapWindowPoints(hCrntWnd, hChildWnd, &pt, 1);
		if (!bRes && (::GetLastError() != ERROR_SUCCESS))
		{
			traceLog << "MapWindowPoints failed in HtmlHelpers::GetIEServerFromScrPt\n";
			break;
		}

		hCrntWnd = hChildWnd;
	}

	return hResult;
}
