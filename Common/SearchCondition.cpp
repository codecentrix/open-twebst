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
#include "SearchCondition.h"


SearchCondition::SearchCondition()
{
	Init();
}


SearchCondition::SearchCondition(BSTR bstrCondition)
{
	ATLASSERT(bstrCondition != NULL);
	Init();
	AddBSTR(bstrCondition);
}


SearchCondition::SearchCondition(LPCSTR szCondition)
{
	USES_CONVERSION;
	ATLASSERT(szCondition != NULL);
	Init();
	AddBSTR(CComBSTR(szCondition));
}


SearchCondition::SearchCondition(LPCWSTR szCondition)
{
	ATLASSERT(szCondition != NULL);
	Init();
	AddBSTR(CComBSTR(szCondition));
}


SearchCondition::SearchCondition(BSTR bstrConditionOne, BSTR bstrConditionTwo)
{
	ATLASSERT(bstrConditionOne != NULL);
	ATLASSERT(bstrConditionTwo != NULL);

	Init();
	AddBSTR(bstrConditionOne);
	AddBSTR(bstrConditionTwo);
}


SearchCondition::SearchCondition(LPCSTR szConditionOne, LPCSTR szConditionTwo)
{
	USES_CONVERSION;
	ATLASSERT(szConditionOne != NULL);
	ATLASSERT(szConditionTwo != NULL);

	Init();
	AddBSTR(CComBSTR(szConditionOne));
	AddBSTR(CComBSTR(szConditionTwo));
}


SearchCondition::SearchCondition(LPCWSTR szConditionOne, LPCWSTR szConditionTwo)
{
	ATLASSERT(szConditionOne != NULL);
	ATLASSERT(szConditionTwo != NULL);

	Init();
	AddBSTR(CComBSTR(szConditionOne));
	AddBSTR(CComBSTR(szConditionTwo));
}


SearchCondition::~SearchCondition()
{
	CleanUp();
}


void SearchCondition::CleanUp()
{
	if (m_pSafeArray != NULL)
	{
		HRESULT hRes = ::SafeArrayDestroy(m_pSafeArray);
		ATLASSERT(SUCCEEDED(hRes));
		m_pSafeArray = NULL;
	}
}


void SearchCondition::Init()
{
	m_nNumberElements = 0;

	SAFEARRAYBOUND rgsabound = { 0, 0 };
	m_pSafeArray = ::SafeArrayCreate(VT_VARIANT, 1, &rgsabound);
	ATLASSERT(m_pSafeArray != NULL);
}


void SearchCondition::AddBSTR(BSTR bstrNew)
{
	// Ignore empty strings.
	if ((NULL == bstrNew) || (L'\0' == bstrNew[0]))
	{
		return;
	}

	// Resize the array.
	CComVariant    varNew(bstrNew);
	SAFEARRAYBOUND rgsabound = { ++m_nNumberElements, 0 };

	HRESULT hRes = ::SafeArrayRedim(m_pSafeArray, &rgsabound);
	ATLASSERT(SUCCEEDED(hRes));

	// Add the string to array.
	long nIndex = { m_nNumberElements - 1 };
	hRes = ::SafeArrayPutElement(m_pSafeArray, &nIndex, &varNew);
	ATLASSERT(SUCCEEDED(hRes));
}


void SearchCondition::operator+=(BSTR bstrCondition)
{
	ATLASSERT(bstrCondition != NULL);
	AddBSTR(bstrCondition);
}


void SearchCondition::operator+=(LPCSTR szCondition)
{
	USES_CONVERSION;
	ATLASSERT(szCondition != NULL);
	AddBSTR(CComBSTR(szCondition));
}


void SearchCondition::operator+=(LPCWSTR szCondition)
{
	ATLASSERT(szCondition != NULL);
	AddBSTR(CComBSTR(szCondition));
}


SearchCondition& SearchCondition::operator=(BSTR bstrCondition)
{
	ATLASSERT(bstrCondition != NULL);
	CleanUp();
	Init();
	AddBSTR(bstrCondition);

	return *this;
}


SearchCondition& SearchCondition::operator=(LPCSTR szCondition)
{
	ATLASSERT(szCondition != NULL);
	CleanUp();
	Init();
	AddBSTR(CComBSTR(szCondition));

	return *this;
}


SearchCondition& SearchCondition::operator=(LPCWSTR szCondition)
{
	ATLASSERT(szCondition != NULL);
	CleanUp();
	Init();
	AddBSTR(CComBSTR(szCondition));

	return *this;
}


// For debug purposes.
CComBSTR SearchCondition::ToBSTR()
{
	if (!m_pSafeArray)
	{
		return L"";
	}

	if (::SafeArrayGetDim(m_pSafeArray) != 1)
	{
		throw CreateException(_T("Wrong number of dimensions in SearchCondition::ToBSTR"));
	}

	long    nLowerBound = 0;
	HRESULT hRes = ::SafeArrayGetLBound(m_pSafeArray, 1, &nLowerBound);
	if (hRes != S_OK)
	{
		throw CreateException(_T("Cannot get lower bound in SearchCondition::ToBSTR"));
	}

	long nUpperBound = 0;
	hRes = ::SafeArrayGetUBound(m_pSafeArray, 1, &nUpperBound);
	if (hRes != S_OK)
	{
		throw CreateException(_T("Cannot get upper bound in SearchCondition::ToBSTR"));
	}

	long     nTokensNumber = nUpperBound - nLowerBound + 1;
	CComBSTR bstrResult    = L"";

	for (int i = 0; i < nTokensNumber; ++i)
	{
		long idx[] = { nLowerBound + i };
		VARIANT var;
		
		::VariantInit(&var);
		hRes = ::SafeArrayGetElement(m_pSafeArray, idx, (void*)&var);
		if (hRes != S_OK)
		{
			::VariantClear(&var);
			throw CreateException(_T("SafeArrayGetElement failed in SearchCondition::ToBSTR"));
		}

		if (var.vt != VT_BSTR)
		{
			::VariantClear(&var);
			throw CreateException(_T("Not a BSTR in the safearray in SearchCondition::ToBSTR"));
		}

		if (bstrResult.Length() == 0)
		{
			bstrResult += L", ";
		}

		bstrResult += var.bstrVal;
	}

	return bstrResult;
}


void SearchCondition::AddMultiCondition(LPCTSTR szMultiCond)
{
	if (!szMultiCond)
	{
		return;
	}

	Common::String sCrntMultiCond = szMultiCond;

	while (TRUE)
	{
		size_t nCommaIndex = sCrntMultiCond.find_first_of(_T(','));
		if (Common::String::npos == nCommaIndex)
		{
			// No more commas.
			this->AddBSTR(CComBSTR(sCrntMultiCond.c_str()));
			break;
		}

		// TODO: un-escape commas. Commas will be represented as &somecode;
		Common::String sCrntCond = sCrntMultiCond.substr(0, nCommaIndex);
		this->AddBSTR(CComBSTR(sCrntCond.c_str()));

		if (sCrntMultiCond.size() == nCommaIndex + 1)
		{
			break;
		}

		sCrntMultiCond = Common::TrimLeftString(sCrntMultiCond.substr(nCommaIndex + 1));
	}
}


BOOL SearchCondition::FindAttribute(LPCTSTR szAttr, CComBSTR& bstrOutVal)
{
	bstrOutVal = L"";

	std::list<Common::DescriptorToken> tokens;
	BOOL bRes = Common::GetDescriptorTokensList(m_pSafeArray, &tokens, NULL);

	if (!bRes)
	{
		return FALSE;
	}

	for (std::list<Common::DescriptorToken>::const_iterator it = tokens.begin(); it != tokens.end(); ++it)
	{
		const Common::String& sAttributeName = it->m_sAttributeName;
		if (!_tcsicmp(sAttributeName.c_str(), szAttr))
		{
			bstrOutVal = it->m_sAttributePattern.c_str();
			return TRUE;
		}
	}

	return FALSE;
}
