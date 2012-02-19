// CatSpyPlugin.cpp : Implementation of CExplorerPlugin
#include "stdafx.h"
#include "DebugServices.h"
#include "Common.h"
#include "HtmlHelpers.h"
#include "ExplorerPlugin.h"


// Spy tree node.
/*struct TreeNode
{
	long					m_nUniqueID;
	CComQIPtr<IHTMLElement> m_spElement;

	TreeNode(long nUniqueID = -1)
	{
		m_nUniqueID = nUniqueID;
	}

	BOOL operator<(const TreeNode& node)
	{
		return (m_nUniqueID < node.m_nUniqueID);
	}
};*/

///////////////////////////////////////////////////////////////////////////////////////////////////////
// ISpyPlugin interface.
//public:
	//STDMETHOD(GetElementFromScreenPos)(LONG nX, LONG nY, IHTMLElement** ppElement);
	//STDMETHOD(GetElementFromCache)    (long nUniqueId, IHTMLElement** ppOutElement);
	//STDMETHOD(ClearCachedInformation) (void);

	// TODO: Code review !!!
	//STDMETHOD(GetElementAttributes)(IHTMLElement* pElement, SAFEARRAY** psa);
	//STDMETHOD(GetElementTree)(IHTMLElement* pElement, SPY_TREE_NODE** pRoot, long* nUniqueID);

//private:
///////////////////////////////////////////////////////////////////////////////////////////////////////
//SPY_TREE_NODE* CreateTreeFromElement(CComQIPtr<IHTMLElement> spElement, long& nCrtId, CComQIPtr<IHTMLElement> spSearchedElement, long* pSearchedId);



//vector<TreeNode> m_treeNodesVector;

STDMETHODIMP CExplorerPlugin::GetElementFromScreenPos(LONG nX, LONG nY, IHTMLElement** ppElement)
{
	// Check the parameters.
	if (NULL == ppElement)
	{
		traceLog << "ppElement is NULL in CExplorerPlugin::GetElementFromScreenPos\n";
		return E_INVALIDARG;
	}

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::GetElementFromScreenPos\n";
		return hRes;
	}

	// Get the html document object.
	CComQIPtr<IDispatch>spDispDoc;
	hRes = spBrws->get_Document(&spDispDoc);

	CComQIPtr<IHTMLDocument2> spDocument = spDispDoc;
	if (spDocument == NULL)
	{
		traceLog << "Can NOT obtain IHTMLDocument in CExplorerPlugin::GetElementFromScreenPos\n";
		return E_FAIL;
	}

	CComQIPtr<IHTMLElement> spElementHit = HtmlHelpers::GetHtmlElementFromScreenPos(spDocument, nX, nY);
	while (TRUE)
	{
		CComQIPtr<IHTMLFrameBase2> spFrameBase = spElementHit;
		if (spFrameBase == NULL)
		{
			break;
		}

		// Get the IHTMLWindow2 object from IHTMLFrameBase2.
		CComQIPtr<IHTMLWindow2> spCurrentWindow;
		hRes = spFrameBase->get_contentWindow(&spCurrentWindow);
		if (spCurrentWindow == NULL)
		{
			traceLog << "IHTMLFrameBase2::get_contentWindow failed with code " << hRes << " in CExplorerPlugin::GetElementFromScreenPos\n";
			return hRes;
		}

		// Get the html document object.
		CComQIPtr<IHTMLDocument2> spCurrentDocument = HtmlHelpers::HtmlWindowToHtmlDocument(spCurrentWindow);
		if (spCurrentDocument == NULL)
		{
			traceLog << "Can not get the html document in CExplorerPlugin::GetElementFromScreenPos\n";
			return E_FAIL;
		}

		spElementHit = HtmlHelpers::GetHtmlElementFromScreenPos(spCurrentDocument, nX, nY);
		if (spElementHit == NULL)
		{
			traceLog << "HtmlHelpers::GetHtmlElementFromScreenPos failed in CExplorerPlugin::GetElementFromScreenPos\n";
			return E_FAIL;
		}
	}

	*ppElement = spElementHit.Detach();
	return S_OK;
}


STDMETHODIMP CExplorerPlugin::GetElementFromCache(long nUniqueId, IHTMLElement** ppOutElement)
{
	if ((NULL == ppOutElement) || (nUniqueId < 0))
	{
		traceLog << "Invalid parameters in CExplorerPlugin::GetElementFromCache\n";
		return E_INVALIDARG;
	}

	TreeNode node(nUniqueId);
	vector<TreeNode>::iterator it = lower_bound(m_treeNodesVector.begin(), m_treeNodesVector.end(), node);

	if (it != m_treeNodesVector.end())
	{
		CComQIPtr<IHTMLElement> spCrtElement = it->m_spElement;

		if (*ppOutElement != NULL)
		{
			(*ppOutElement)->Release();
		}

		*ppOutElement = spCrtElement.Detach();
	}
	else
	{
		traceLog << "The node with id: " << nUniqueId << " was not found in CExplorerPlugin::GetElementFromCache\n";
		return S_FALSE;
	}

	return S_OK;
}


STDMETHODIMP CExplorerPlugin::ClearCachedInformation(void)
{
	m_treeNodesVector.clear();
	return S_OK;
}














/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
// TODO: Code review !!!


#define MAX_CHARACTER_COUNT	10


STDMETHODIMP CExplorerPlugin::GetElementAttributes(IHTMLElement* pElement, SAFEARRAY** psa)
{
	//Get selected html element attributes.
	CComQIPtr<IHTMLDOMNode> spDomNode = pElement;
	//ASSERT(spDomNode != NULL);
	if (spDomNode == NULL)
	{
		traceLog << "Can NOT obtain IHTMLDOMNode from IHTMLElement in CExplorerPlugin::GetElementAttributes\n";
		return S_FALSE;
	}

	// Get an IDispatch pointer to the collection of attributes.
	CComQIPtr<IDispatch> spDispAttributes;
	spDomNode->get_attributes(&spDispAttributes);
	//ASSERT(spDispAttributes != NULL);
	if (spDispAttributes == NULL)
	{
		traceLog << "Can NOT obtain attributes collection in CExplorerPlugin::GetElementAttributes\n";
		return S_FALSE;
	}

	// Query for IHTMLAttributeCollection interface to parse the attributes collection.
	CComQIPtr<IHTMLAttributeCollection> spCollection = spDispAttributes;
	//ASSERT(spCollection != NULL);
	if (spCollection == NULL)
	{
		traceLog << "Can NOT obtain IHTMLAttributeCollection in CExplorerPlugin::GetElementAttributes\n";
		return S_FALSE;
	}

	long numChildren = 0;
	spCollection->get_length(&numChildren);

	//Intermediary vector with attributes names and values.
	vector<CComBSTR> vecAttributes;


	//Get the text of the element. This should be done by a special algorithm.
	//For now just get the html outer text.
	CComBSTR bstrAllText;
	pElement->get_outerText(&bstrAllText);
	if ((bstrAllText != NULL) && (bstrAllText.Length() > 0))
	{
		//Maybe we should truncate the text. Don't know yet.
		USES_CONVERSION;
		String sAllText(W2T(bstrAllText));
		String sText;

		/*if (sAllText.length() > MAX_CHARACTER_COUNT)
		{
			sText = sAllText.substr(0, MAX_CHARACTER_COUNT);
		}
		else
		{
			sText = sAllText;
		}*/

		size_t nFirstEnter = sAllText.find_first_of(_T('\r'));
		if (nFirstEnter != String::npos)
		{
			sText = sAllText.substr(0, nFirstEnter);
		}
		else
		{
			sText = sAllText;
		}

		if (!sText.empty())
		{
			vecAttributes.push_back(CComBSTR("Text"));
			vecAttributes.push_back(CComBSTR(sText.c_str()) /*bstrAllText*/);
		}
	}
	///////////////////////////////////////


	//Browser all attributes of the currently selected html element.
	for (long i = 0; i < numChildren; i++)
	{
		// Get the current attribute element.
		CComVariant index(i);
		CComPtr<IDispatch> spItemDisp;
		spCollection->item(&index, &spItemDisp);
		if (spItemDisp == NULL)
		{
			continue;
		}

		CComQIPtr<IHTMLDOMAttribute> spAttribute = spItemDisp;
		if (spAttribute == NULL)
		{
			continue;
		}
		
		//Get attribute name and value.
		CComBSTR bstrNodeName;
		spAttribute->get_nodeName(&bstrNodeName);
		//ASSERT((bstrNodeName.GetBSTR() != NULL) && (bstrNodeName.length() > 0));
		if ((bstrNodeName == NULL) || (bstrNodeName.Length() == 0))
		{
			// If attribute name is not valid go to the next attribute in the collection.
			continue;
		}

		CComVariant varNodeValue;
		spAttribute->get_nodeValue(&varNodeValue);

		HRESULT hRes = S_OK;
		if (varNodeValue.vt != VT_BSTR)
		{
			// Try to convert the attribute value to a BSTR.
			hRes = ::VariantChangeType(&varNodeValue, &varNodeValue,
				VARIANT_ALPHABOOL | VARIANT_NOUSEROVERRIDE | VARIANT_LOCALBOOL, VT_BSTR);
		}

		if (S_OK == hRes)
		{
			CComBSTR bstrNodeValue;
			//ASSERT(VT_BSTR == varNodeValue.vt);
			if (VT_BSTR == varNodeValue.vt)
			{
				bstrNodeValue = varNodeValue.bstrVal;
			}

			if ((bstrNodeValue != NULL) && (bstrNodeValue.Length() > 0))
			{
				// Add (name, value) pair to our collection only if name and value are valid.
				vecAttributes.push_back(bstrNodeName);
				vecAttributes.push_back(bstrNodeValue);
			}
		}
	}

	//Create a new safearray to be filled with attributes.
	SAFEARRAYBOUND bounds;
	bounds.cElements = vecAttributes.size();
	bounds.lLbound = 0;
	*psa = ::SafeArrayCreate(VT_BSTR, 1, &bounds);

	if (NULL == *psa)
	{
		traceLog << "Can NOT create a SAFEARRAY in CExplorerPlugin::GetElementAttributes\n";
		return S_FALSE;
	}

	//Populate the new created safearray with values from the temporary list.
	for (long i = 0; i < bounds.cElements; ++i)
	{
		::SafeArrayPutElement(*psa, &i, vecAttributes[i]);
	}

	return S_OK;
}


SPY_TREE_NODE* CreateNewNode()
{
	SPY_TREE_NODE* pLocalRoot = (SPY_TREE_NODE*)::CoTaskMemAlloc(sizeof(SPY_TREE_NODE));
	pLocalRoot->m_nChildrenCount  = 0;
	pLocalRoot->m_pChildrenVector = NULL;
	return pLocalRoot;
}


SPY_TREE_NODE* CExplorerPlugin::CreateTreeFromElement(CComQIPtr<IHTMLElement> spElement, long& nCrtId,
							CComQIPtr<IHTMLElement> spSearchedElement, long* pSearchedId)
{
	if (spElement == NULL)
	{
		return NULL;
	}

	// Get the Text and Tag of the current html element.
	CComBSTR bstrOuterText;
	spElement->get_outerText(&bstrOuterText);
	CComBSTR bstrTagName;
	spElement->get_tagName(&bstrTagName);

	//Get at most MAX_CHARACTER_COUNT characters of element Text.
	String sNodeText;
	if (bstrOuterText != NULL)
	{
		USES_CONVERSION;
		String sOutertext(W2T(bstrOuterText));
		if (sOutertext.length() > MAX_CHARACTER_COUNT)
		{
			sNodeText = sOutertext.substr(0, MAX_CHARACTER_COUNT);
			sNodeText += _T("...");
		}
		else
		{
			sNodeText = sOutertext;
		}

		size_t nFirstEnter = sNodeText.find_first_of(_T('\r'));
		if (nFirstEnter != String::npos)
		{
			sNodeText = sNodeText.substr(0, nFirstEnter);
		}
	}
	sNodeText += _T(" (");
	//Get at most MAX_CHARACTER_COUNT from element Tag.
	if (bstrTagName != NULL)
	{
		/*String sTagName((LPTSTR)bstrTagName);
		if (sTagName.length() > MAX_CHARACTER_COUNT)
		{
			sNodeText += sTagName.substr(0, MAX_CHARACTER_COUNT);
		}
		else
		{
			sNodeText += sTagName;
		}*/

		USES_CONVERSION;
		sNodeText += W2T(bstrTagName);
	}
	sNodeText += _T(")");

	// Create a new tree node and set text and unique id.
	SPY_TREE_NODE* pLocalRoot = ::CreateNewNode();
	pLocalRoot->m_bstrNodeName = CComBSTR(sNodeText.c_str()).Detach();
	nCrtId++;
	pLocalRoot->m_nUniqueID = nCrtId;

	// Append the current html pointer to the cache vector of pointers.
	// This list of pointers is created so we don't have to marshall all html elements
	// across the processes boundaries.
	TreeNode cache_node;
	cache_node.m_nUniqueID = nCrtId;
	cache_node.m_spElement = spElement;
	m_treeNodesVector.push_back(cache_node);

	//If the searched element hasn't been found so far check if the current element is the one.
	if ((*pSearchedId) < 0)
	{
		if (spElement.IsEqualObject(spSearchedElement))
		{
			*pSearchedId = nCrtId;
		}
	}

	//See if the current element is a html frame actually.
	CComQIPtr<IHTMLFrameBase2> spFrameBase = spElement;
	if (spFrameBase != NULL)
	{
		CComQIPtr<IHTMLWindow2> spWindow2;
		spFrameBase->get_contentWindow(&spWindow2);
		if (spWindow2 != NULL)
		{
			CComQIPtr<IHTMLDocument2> spDocument = HtmlHelpers::HtmlWindowToHtmlDocument(spWindow2);
			if (spDocument != NULL)
			{
				CComQIPtr<IHTMLElement> spBody;
				spDocument->get_body(&spBody);
				if (spBody != NULL)
				{
					pLocalRoot->m_nChildrenCount = 1;
					pLocalRoot->m_pChildrenVector = (SPY_TREE_NODE**)::CoTaskMemAlloc(sizeof(SPY_TREE_NODE*));
					pLocalRoot->m_pChildrenVector[0] = CreateTreeFromElement(spBody, nCrtId,
													spSearchedElement, pSearchedId);
					return pLocalRoot;
				}
			}
		}
	}

	// Get the children collection.
	CComQIPtr<IDispatch> spChildrenDisp;
	spElement->get_children(&spChildrenDisp);
	CComQIPtr<IHTMLElementCollection> spChildren = spChildrenDisp;
	if (spChildren == NULL)
	{
		return pLocalRoot;
	}

	long lChildrenCount = 0;
	spChildren->get_length(&lChildrenCount);
	pLocalRoot->m_nChildrenCount = lChildrenCount;

	if (lChildrenCount <= 0)
	{
		return pLocalRoot;
	}

	pLocalRoot->m_pChildrenVector = (SPY_TREE_NODE**)::CoTaskMemAlloc(lChildrenCount * sizeof(SPY_TREE_NODE*));

	// Iterate throw the children collection.
	for (long i = 0; i < lChildrenCount; ++i)
	{
		CComQIPtr<IDispatch> spItemDisp;
		spChildren->item(CComVariant(i), CComVariant(), &spItemDisp);
		CComQIPtr<IHTMLElement> spCurrentElement = spItemDisp;

		//Recursively call for each child, from left to right (inorder).
		pLocalRoot->m_pChildrenVector[i] = CreateTreeFromElement(spCurrentElement, nCrtId,
														spSearchedElement, pSearchedId);
	}

	return pLocalRoot;
}


STDMETHODIMP CExplorerPlugin::GetElementTree(IHTMLElement* pElement, SPY_TREE_NODE** pRoot, 
											 long* nUniqueID)
{
	*pRoot = NULL;

	// Get the IWebBrowser2 pointer.
	CComQIPtr<IWebBrowser2> spBrws;
	HRESULT hRes = GetSite(IID_IWebBrowser2, (void**)&spBrws);
	if (spBrws == NULL)
	{
		traceLog << "Can NOT obtain IWebBrowser2 in CExplorerPlugin::GetElementTree\n";
		return hRes;
	}

	CComQIPtr<IHTMLDocument2>	spDocument;
	CComQIPtr<IDispatch>		spDispDoc;

	hRes = spBrws->get_Document(&spDispDoc);
	spDocument = spDispDoc;

	if (spDocument == NULL)
	{
		traceLog << "Can NOT obtain IHTMLDocument in CExplorerPlugin::GetElementTree\n";
		return hRes;
	}

	CComQIPtr<IHTMLElement> spBody;
	hRes = spDocument->get_body(&spBody);
	if (spBody == NULL)
	{
		traceLog << "Can NOT obtain body element in CExplorerPlugin::GetElementTree\n";
		return hRes;
	}

	//Clear all html elementes stored.
	m_treeNodesVector.clear();

	long nCrtId			= -1;
	long nSearchedId	= -1;
	*pRoot = CreateTreeFromElement(spBody, nCrtId, CComQIPtr<IHTMLElement>(pElement), 
																	&nSearchedId);

	*nUniqueID = nSearchedId;

	return S_OK;
}

