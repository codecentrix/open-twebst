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
#include "FindInTimeout.h"


struct SearchContext;


// Base class for code reuse only.
class BaseLibObject : public IAutoClosePopups, public ISleeper
{
public:
	// Public methods.
	BaseLibObject(const IID& clsid) : m_clsid(clsid), m_nPluginThreadID(0)
	{
	}

	HRESULT SetComErrorMessage    (UINT nId, DWORD dwHelpID);
	HRESULT SetComErrorMessage    (UINT nIdPattern, LPCTSTR szFunctionName, DWORD dwHelpID);
	void    SetPlugin             (IExplorerPlugin* pPlugin);
	void    SetCore               (ICore* pCore);
	
	virtual void CloseBrowserPopups();
	virtual BOOL Sleep(DWORD dwMiliseconds);
	static BOOL InitCloseBrowserPopups();


protected:
	virtual BOOL IsValidState(UINT nPatternID, LPCTSTR szMethodName, DWORD dwHelpID);

	void    SetLastErrorCode     (LONG nErr);
	BOOL    GetPluginThreadID    (LONG& nThreadID);
	HRESULT FindInContainer      (IUnknown* pContainer, SearchContext& sc, BSTR bstrTag, SAFEARRAY* pVarArgs, IUnknown** ppResult);
	HRESULT CreateElementObject  (IUnknown* pHtmlElement, IElement** ppElement);
	HRESULT CreateFrameObject    (IUnknown* pHtmlFrame,   IFrame**   ppFrame);
	HRESULT CreateElemListObject (IUnknown* pLocalElemCollection, IElementList** ppElementList);
	static BOOL  LoadAutoCloseSettings(const String& sXmlFilePath);

private:
	HRESULT Find(IUnknown* pContainer, SearchContext& sc, BSTR bstrTag, SAFEARRAY* pVarArgs, void** ppResult);

protected:
	CComQIPtr<IExplorerPlugin> m_spPlugin;
	CComQIPtr<ICore>           m_spCore;
	const IID&                 m_clsid;

	// Auto-close popups.
	LONG  m_nPluginThreadID;
	static CAtlMap<CString, int> s_autoClosePopupsMap;
	static BOOL                  s_bPopupsMapInit;
};
