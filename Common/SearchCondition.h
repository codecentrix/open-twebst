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


class SearchCondition
{
public:
	 SearchCondition();
	 SearchCondition(BSTR    bstrCondition);
	 SearchCondition(LPCSTR  szCondition);
	 SearchCondition(LPCWSTR szCondition);

	 SearchCondition(BSTR    bstrConditionOne, BSTR    bstrConditionTwo);
	 SearchCondition(LPCSTR  szConditionOne,   LPCSTR  szConditionTwo);
	 SearchCondition(LPCWSTR szConditionOne,   LPCWSTR szConditionTwo);

	~SearchCondition();

	void operator+=(BSTR    bstrCondition);
	void operator+=(LPCSTR  szCondition);
	void operator+=(LPCWSTR szCondition);
	operator LPSAFEARRAY() { return m_pSafeArray; }
	SearchCondition& operator=(BSTR    bstrCondition);
	SearchCondition& operator=(LPCSTR  szCondition);
	SearchCondition& operator=(LPCWSTR szCondition);

	CComBSTR ToBSTR           ();
	void     AddMultiCondition(LPCTSTR szMultiCond);
	BOOL     FindAttribute    (LPCTSTR szAttr, CComBSTR& bstrOutVal);


private:
	// Avoid duplicating search conditions.
	SearchCondition& operator=(const SearchCondition&);
	SearchCondition(const SearchCondition&);

	void Init();
	void AddBSTR(BSTR bstrNew);
	void CleanUp();

private:
	SAFEARRAY* m_pSafeArray;
	long       m_nNumberElements;
};
