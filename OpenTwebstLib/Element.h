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
#include "resource.h"       // main symbols
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "CatCLSLib.h"
#include "BaseLibObject.h"
struct SelectCallContext;
struct ClickCallContext;


class ATL_NO_VTABLE CElement : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CElement>,
	public ISupportErrorInfo,
	public IDispatchImpl<IElement, &IID_IElement, &LIBID_OpenTwebstLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public BaseLibObject
{
public:
	CElement() : BaseLibObject(GetObjectCLSID())
	{
	}

	void SetHtmlElement(IHTMLElement* pHtmlElement);


DECLARE_REGISTRY_RESOURCEID(IDR_ELEMENT)


BEGIN_COM_MAP(CElement)
	COM_INTERFACE_ENTRY(IElement)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	// IElement
	STDMETHOD(get_uiName)                (BSTR* pVal);
	STDMETHOD(get_parentBrowser)         (IBrowser**     pVal);
	STDMETHOD(get_parentFrame)           (IFrame**       pVal);
	STDMETHOD(get_nextSiblingElement)    (IElement**     pVal);
	STDMETHOD(get_previousSiblingElement)(IElement**     pVal);
	STDMETHOD(get_parentElement)         (IElement**     pVal);
	STDMETHOD(get_nativeElement)         (IHTMLElement** pVal);
	STDMETHOD(get_core)                  (ICore**        pVal);
	STDMETHOD(FindElement)               (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(FindChildElement)          (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(FindAllElements)           (BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList);
	STDMETHOD(FindChildrenElements)      (BSTR bstrTag, BSTR bstrCond, IElementList** ppElementList);
	STDMETHOD(Click)                     ();
	STDMETHOD(RightClick)                ();
	STDMETHOD(InputText)                 (BSTR bstrText);
	STDMETHOD(Select)                    (VARIANT vItems);
	STDMETHOD(AddSelection)              (VARIANT vItems);
	STDMETHOD(SelectRange)               (VARIANT vStart, VARIANT vEnd);
	STDMETHOD(AddSelectionRange)         (VARIANT vStart, VARIANT vEnd);
	STDMETHOD(ClearSelection)            (void);
	STDMETHOD(Highlight)                 (void);
	STDMETHOD(GetAttribute)              (BSTR bstrAttrName, VARIANT* pVal);
	STDMETHOD(SetAttribute)              (BSTR bstrAttrName, VARIANT varAttrValue);
	STDMETHOD(RemoveAttribute)           (BSTR bstrAttrName);
	STDMETHOD(get_tagName)               (BSTR* pVal);
	STDMETHOD(FindParentElement)         (BSTR bstrTag, BSTR bstrCond, IElement** ppElement);
	STDMETHOD(get_isChecked)             (VARIANT_BOOL*  pIsChecked);
	STDMETHOD(Check)                     (void);
	STDMETHOD(Uncheck)                   (void);
	STDMETHOD(get_selectedOption)        (IElement**     ppSelectedOption);
	STDMETHOD(GetAllSelectedOptions)     (IElementList** ppSelectedOptionsList);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease() 
	{
	}


protected:
	virtual BOOL IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID);

private:
	enum HTML_EDIT_BOX_TYPE
	{
		HTML_NOT_EDIT_BOX,
		HTML_TEXT_EDIT_BOX,
		HTML_PASSWORD_EDIT_BOX,
		HTML_MULTILINE_EDIT_BOX,
		HTML_FILE_EDIT_BOX,
	};

private:
	BOOL    DrawRectangleInWindow (HWND hWnd, long nLeft, long nRight, long nTop, long nBottom);
	HRESULT Click                 (const ClickCallContext& clickCtx, BOOL bClickFileInputButton, BOOL* pbPostedClick = NULL);
	HWND    GetTopWnd             ();
	HRESULT PressOpenBtnInFileDlg (HWND hDlgWnd);
	HRESULT InputTextInFileCtrlIE8(BSTR bstrText);
	HRESULT InputTextInFileDlg    (BSTR bstrText, HWND hIeWnd);
	HRESULT InputTextInFileDlg    (HWND hDlgWnd,  BSTR bstrText);
	HRESULT IeInputText           (BSTR bstrText, BOOL bIsInputFileType, VARIANT_BOOL vbAsync);
	HRESULT HardwareInputText     (BSTR bstrText, BOOL bIsInputFileType);
	HRESULT InputTextInElement    (IHTMLElement* pElement, BSTR bstrText);
	HRESULT GetParentWindow       (IHTMLElement* pElement, IHTMLWindow2** ppElement);
	HRESULT IsHtmlEditBox         (IHTMLElement* pElement, HTML_EDIT_BOX_TYPE* pEditBoxType);
	HRESULT Select                (const VARIANT& vStartItems, const VARIANT& vEndItems, const SelectCallContext& context);
	BOOL    GetElemValue          (CComQIPtr<IHTMLElement> spHtmlElement, BSTR* pBstrInitialValue);
	BOOL    IsWindowActive        (HWND hIEWnd);
	BOOL    IsCheckable           (CComQIPtr<IHTMLElement> spElement, BOOL bRadioIsNotCheckable = FALSE);
	HRESULT GetHandlerAttrText    (BSTR bstrAttrName, VARIANT* pVal);

	static BOOL FindFileDlgCallback(HWND hWnd, void* pObj);
	static BOOL FindOpenBtnCallback(HWND hWnd, void* pObj);

private:
	CComQIPtr<IHTMLElement> m_spHtmlElement;
};
