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
#include "..\BrowserPlugin\BrowserPlugin.h"
#include "Core.h"


class IAutoClosePopups
{
public:
	virtual void CloseBrowserPopups() = 0;
};


class ISleeper
{
public:
	virtual BOOL Sleep(DWORD dwMiliseconds, BOOL bDispatchMsg = FALSE) = 0;
};


//////////////////////////////////////////////////////////////////////////////////////////////
class FinderInTimeout
{
public:
	FinderInTimeout(CComQIPtr<IExplorerPlugin> spPlugin,
	                LONG nSearchFlags,
                    SAFEARRAY* pVarArgs,
					ICore* pCore,
					CComQIPtr<IUnknown, &IID_IUnknown> spContainer)
                    : m_spPlugin(spPlugin), m_nSearchFlags(nSearchFlags), m_pVarArgs(pVarArgs), m_spCore(pCore), m_spContainer(spContainer)
	{
		ATLASSERT(spPlugin != NULL);
		ATLASSERT(pCore    != NULL);
	}

	virtual BOOL Find() = 0;

	// Returns FALSE if timeout.
	virtual BOOL WaitToLoad(DWORD nLoadTimeout, ISleeper* pSleeper)
	{
		UNREFERENCED_PARAMETER(nLoadTimeout);
		return TRUE; // No timeout.
	}

	virtual ~FinderInTimeout() = 0
	{
	}

	BOOL IsBrowserLoading();
	static BOOL Find(FinderInTimeout*  pFinder,
                     ULONG             nSearchTimeout,
                     ULONG             nLoadTimeout,
                     VARIANT_BOOL      vbLoadTimeoutIsErr,
					 IAutoClosePopups* pAutoClosePopups,
					 ISleeper*         pSleeper);

protected:
	CComQIPtr<IExplorerPlugin>         m_spPlugin;
	LONG                               m_nSearchFlags;
	SAFEARRAY*                         m_pVarArgs;
	CComQIPtr<ICore>                   m_spCore;
	CComQIPtr<IUnknown, &IID_IUnknown> m_spContainer;
};


//////////////////////////////////////////////////////////////////////////////////////////////
class FindObjectInContainer : public FinderInTimeout
{
public:
	FindObjectInContainer(CComQIPtr<IUnknown, &IID_IUnknown> spContainer,
	                      CComQIPtr<IExplorerPlugin> spPlugin,
                          LONG nSearchFlags, BSTR bstrTag,
						  SAFEARRAY* pVarArgs,
						  ICore* pCore) :
		 m_bstrTag(bstrTag), FinderInTimeout(spPlugin, nSearchFlags, pVarArgs, pCore, spContainer)
	{
		ATLASSERT((bstrTag != NULL) ||
		          ((NULL == bstrTag) &&
				  ((nSearchFlags & Common::SEARCH_FRAME)  || (nSearchFlags & Common::SEARCH_HTML_DIALOG) ||
				   (nSearchFlags & SEARCH_MODAL_HTML_DLG) || (nSearchFlags & Common::SEARCH_SELECTED_OPTION))));
	}

	CComQIPtr<IUnknown, &IID_IUnknown> GetResultObject() const
	{
		return m_spResult;
	}

	virtual BOOL Find();
	virtual BOOL WaitToLoad(DWORD nLoadTimeout, ISleeper*  pSleeper);


private:
	CComBSTR                           m_bstrTag;
	CComQIPtr<IUnknown, &IID_IUnknown> m_spResult;
};



//////////////////////////////////////////////////////////////////////////////////////////////
class FindFramesInContainer : public FinderInTimeout
{
public:
	FindFramesInContainer(CComQIPtr<IUnknown, &IID_IUnknown> spContainer,
	                      CComQIPtr<IExplorerPlugin> spPlugin,
                          LONG nSearchFlags, SAFEARRAY* pVarArgs,
						  ICore* pCore) :
		 FinderInTimeout(spPlugin, nSearchFlags, pVarArgs, pCore, spContainer)
	{
		m_ppFrames = NULL;
	}

	virtual BOOL Find();
	virtual ~FindFramesInContainer();
	void Detach(IHTMLWindow2*** pppFrames, LONG* pSize);


private:
	LONG                               m_nCollectionSize;
	IHTMLWindow2**                     m_ppFrames;
};



//////////////////////////////////////////////////////////////////////////////////////////////
class SelectOptionsInContainer : public FinderInTimeout
{
public:
	SelectOptionsInContainer
		(
			IUnknown*                  pContainer,
			CComQIPtr<IExplorerPlugin> spPlugin,
			const VARIANT&             vStartItems,
			const VARIANT&             vEndItems,
			LONG                       nFlags,
			CComQIPtr<ICore>           spCore
		) : FinderInTimeout(spPlugin, nFlags, NULL, spCore, pContainer)
	{
		m_vStartItems = vStartItems;
		m_vEndItems   = vEndItems;
		m_bExecuted   = FALSE;
	}

	virtual BOOL Find();
	BOOL IsSelectionExecuted()
	{
		return m_bExecuted;
	}

private:
	CComVariant m_vStartItems;
	CComVariant m_vEndItems;
	BOOL        m_bExecuted;
};
