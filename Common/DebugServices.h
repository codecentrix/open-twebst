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


// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the DBGSERV_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// DBGSERV_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef DBGSERV_EXPORTS
#define DBGSERV_API __declspec(dllexport)
#else
#define DBGSERV_API __declspec(dllimport)
#endif


// Types, functions, classes used for debugging purposes inside the DebugServices namespace.
#pragma once
#include "Common.h"
#include "Exceptions.h"

using namespace Common;
using namespace ExceptionServices;


#ifndef DISABLE_TRACE_SERVICE
namespace DebugServices
{
	//////////////////////////////////////////
	class TraceService
	{
	public:
		DBGSERV_API static TraceService& GetGlobalTraceService(int nCodeLineNumber = 0, const CHAR* szSourceFileName = "");
		DBGSERV_API const  TraceService& operator << (LPCSTR) const;

		inline static void Init();
		void ThreadUnInit();
		~TraceService();

		// Template << operator for any other type that could be converted to string by ostringstream.
		template <typename T>
			inline const TraceService& operator << (T tElement) const;

		inline const TraceService& operator << (HRESULT hRes                        ) const;
		inline const TraceService& operator << (CHAR str[]                          ) const;
		inline const TraceService& operator << (const Exception&                    ) const;
		inline const TraceService& operator << (const RegistryException&            ) const;
		inline const TraceService& operator << (const InvalidParamException&        ) const;
		inline const TraceService& operator << (const BrowserDisconnectedException& ) const;
		inline const TraceService& operator << (const OperationNotAllowedException& ) const;
		inline const TraceService& operator << (const IndexOutOfBoundException&     ) const;
		inline const TraceService& operator << (const String&                       ) const;
		inline const TraceService& operator << (LPCWSTR                             ) const;
		inline const TraceService& operator << (WCHAR wstr[]                        ) const;
		inline const TraceService& operator << (const CComBSTR&                     ) const;
		inline const TraceService& operator << (WCHAR                               ) const;

	private:
		DBGSERV_API TraceService();
		TraceService(const TraceService& );

		int         GetCodeLineNumber() const;
		const CHAR* GetSourceFileName() const;
		BOOL        SetCodeLineNumber(int nCodeLineNumber) const;
		BOOL        SetSourceFileName(const CHAR*) const;

	private:
		struct TlsData
		{
			TlsData() : m_nCodeLineNumber(0), m_szSourceFileName(NULL)
			{
			}

			int         m_nCodeLineNumber;
			const CHAR* m_szSourceFileName;
		};

	private:
		DWORD       m_dwTlsIndex;
		BOOL        m_bLogEnabled;
		std::string m_sLogFile;
	};



	// Inline methods implementation.
	inline void TraceService::Init()
	{
		GetGlobalTraceService();
	}

	template <typename T>
	inline const TraceService& TraceService::operator << (typename T tElement) const
	{
		std::ostringstream outputStream;

		outputStream << tElement;
		return operator << (outputStream.str().c_str());
	}

	inline const TraceService& TraceService::operator << (HRESULT hRes) const
	{
		std::ostringstream outputStream;

		outputStream << std::hex << hRes;
		return operator << (outputStream.str().c_str());
	}

	inline const TraceService& TraceService::operator << (const Exception& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const RegistryException& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const BrowserDisconnectedException& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const InvalidParamException& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const OperationNotAllowedException& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const IndexOutOfBoundException& except) const
	{
		return operator << (except.ToString());
	}

	inline const TraceService& TraceService::operator << (const String& str) const
	{
		return operator << (str.c_str());
	}

	inline const TraceService& TraceService::operator << (const LPCWSTR lpcwstr) const
	{
		USES_CONVERSION;
		return operator << (W2A(lpcwstr));
	}

	inline const TraceService& TraceService::operator << (WCHAR wch) const
	{
		USES_CONVERSION;
		WCHAR wstr[2] = { wch, L'\0' };

		return operator << (W2A(wstr));
	}

	inline const TraceService& TraceService::operator << (WCHAR wstr[]) const
	{
		return operator << (static_cast<LPCWSTR>(wstr));
	}

	inline const TraceService& TraceService::operator << (CHAR str[]) const
	{
		return operator << (static_cast<LPCSTR>(str));
	}
}


// Global traceLog macro definition
#define traceLog DebugServices::TraceService::GetGlobalTraceService(__LINE__, __FILE__)



#else
namespace DebugServices
{
	class TraceService
	{
	public:

		// Template << operator for any other type that could be converted to string by ostringstream.
		template <typename T>
			inline const TraceService& operator << (T tElement) const
		{
			return *this;
		}
	};

	const TraceService PER_FILE_TRACE_SERVICE_OBJECT;
}

#define traceLog DebugServices::PER_FILE_TRACE_SERVICE_OBJECT

#endif //DISABLE_TRACE_SERVICE
