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
#include "Registry.h"
using namespace ExceptionServices;

namespace DebugServices
{
	const TCHAR REG_TRACE_KEY_NAME[]        = _T("Software\\CodeCentrix\\OpenTwebst");
	const TCHAR REG_TRACE_FLAG_VALUE_NAME[] = _T("EnableTrace");
	const TCHAR REG_TRACE_FILE_NAME[]       = _T("TraceFile");

	////////////////////////////////////////////
	DBGSERV_API TraceService& TraceService::GetGlobalTraceService(int nCodeLineNumber, const CHAR* szSourceFileName)
	{
		static DebugServices::TraceService traceLogObject;

		traceLogObject.SetCodeLineNumber(nCodeLineNumber);
		traceLogObject.SetSourceFileName(szSourceFileName);
		return traceLogObject;
	}


	// Constructor.
	DBGSERV_API TraceService::TraceService()
	{
		try
		{
			// Read the trace file name and the log enable/disabled flag from the registry.
			DWORD dwTraceEnabled = Registry::RegGetDWORDValue(HKEY_CURRENT_USER, REG_TRACE_KEY_NAME, REG_TRACE_FLAG_VALUE_NAME);
			m_bLogEnabled = (dwTraceEnabled != 0);

			if (m_bLogEnabled)
			{
				String sLogFile  = Registry::RegGetStringValue(HKEY_CURRENT_USER, REG_TRACE_KEY_NAME, REG_TRACE_FILE_NAME);

				USES_CONVERSION;
				m_sLogFile = T2A(const_cast<TCHAR*>(sLogFile.c_str()));	// Log file is always ASCII.
			}
		}
		catch (const RegistryException& regException)
		{
#ifdef _DEBUG
			::OutputDebugString(_T("Registry exception while initializing the trace system\n"));
			::OutputDebugString(regException.ToString().c_str());
			::OutputDebugString(_T("\n"));
#else
			// Avoid "unreferenced local variable" warning.
			regException;
#endif

			// The traces are disabled by default.
			m_bLogEnabled = FALSE;
		}

		// Alloc a thread local index.
		m_dwTlsIndex = ::TlsAlloc();
	}

	TraceService::~TraceService()
	{
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			::TlsFree(m_dwTlsIndex);
			m_dwTlsIndex = TLS_OUT_OF_INDEXES;
		}
	}

	DBGSERV_API const TraceService& TraceService::operator << (LPCSTR lpcstr) const
	{
		const CHAR* szSourceFileName = GetSourceFileName();
		int         nCodeLineNumber  = GetCodeLineNumber();

#ifdef _DEBUG
		// Always use the OutputDebugString in debug version.
		if (nCodeLineNumber != 0)
		{
			std::ostringstream outputStream;
			outputStream << szSourceFileName << "(" << nCodeLineNumber << ") ";

			::OutputDebugStringA(outputStream.str().c_str());
			if (!m_bLogEnabled)
			{
				nCodeLineNumber = 0;
				SetCodeLineNumber(0);
			}
		}

		::OutputDebugStringA(lpcstr);
#endif

		if (m_bLogEnabled)
		{
			// Compute the trace file name. The name includes the thread id. There will be
			// on tace log file per thread.
			std::ostringstream oFileNameStream;
			oFileNameStream << m_sLogFile << ::GetCurrentThreadId() << ".txt";

			// Open the log file. Log file is always ASCII.
			std::ofstream traceFile(oFileNameStream.str().c_str(), std::ios_base::app);
			if (traceFile.is_open())
			{
				if (nCodeLineNumber != 0)
				{
					traceFile << szSourceFileName << "(" << nCodeLineNumber << ") ";
				}

				traceFile << lpcstr;
				traceFile.close();
			}

			SetCodeLineNumber(0);
		}

		return *this;
	}


	// Returns 0 on error.
	int TraceService::GetCodeLineNumber() const
	{
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			LPVOID         lpData   = ::TlsGetValue(m_dwTlsIndex);
			const TlsData* pTlsData = static_cast<const TlsData*>(lpData);
			
			if (pTlsData != NULL)
			{
				return pTlsData->m_nCodeLineNumber;
			}
		}

		return 0;
	}


	// Returns NULL on error.
	const CHAR* TraceService::GetSourceFileName() const
	{
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			LPVOID         lpData   = ::TlsGetValue(m_dwTlsIndex);
			const TlsData* pTlsData = static_cast<const TlsData*>(lpData);
			
			if (pTlsData != NULL)
			{
				return pTlsData->m_szSourceFileName;
			}
		}

		return NULL;
	}


	// Returns FALSE on error.
	BOOL TraceService::SetCodeLineNumber(int nCodeLineNumber) const
	{
		_ASSERTE(nCodeLineNumber >= 0);
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			LPVOID   lpData   = ::TlsGetValue(m_dwTlsIndex);
			TlsData* pTlsData = static_cast<TlsData*>(lpData);
			
			if (NULL == pTlsData)
			{
				pTlsData = new TlsData();

				BOOL bRes = ::TlsSetValue(m_dwTlsIndex, static_cast<LPVOID>(pTlsData));
				if (FALSE == bRes)
				{
					return FALSE;
				}
			}

			pTlsData->m_nCodeLineNumber = nCodeLineNumber;
			return TRUE;
		}

		return FALSE;
	}


	// Returns FALSE on error.
	BOOL TraceService::SetSourceFileName(const CHAR* szSourceFileName) const
	{
		_ASSERTE(szSourceFileName != NULL);
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			LPVOID   lpData   = ::TlsGetValue(m_dwTlsIndex);
			TlsData* pTlsData = static_cast<TlsData*>(lpData);
			
			if (NULL == pTlsData)
			{
				pTlsData = new TlsData();

				BOOL bRes = ::TlsSetValue(m_dwTlsIndex, static_cast<LPVOID>(pTlsData));
				if (FALSE == bRes)
				{
					return FALSE;
				}
			}

			pTlsData->m_szSourceFileName = szSourceFileName;
			return TRUE;
		}

		return FALSE;
	}


	void TraceService::ThreadUnInit()
	{
		if (m_dwTlsIndex != TLS_OUT_OF_INDEXES)
		{
			LPVOID         lpData   = ::TlsGetValue(m_dwTlsIndex);
			const TlsData* pTlsData = static_cast<const TlsData*>(lpData);
			delete pTlsData;
		}
	}
}
