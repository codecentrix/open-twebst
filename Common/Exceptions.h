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

// This header contains the base class Exception. All the exceptions used
// in project will be ultimately from Exception under the namespace ErrorServices.
#pragma once
#include "Common.h"


namespace ExceptionServices
{
	using namespace Common;

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Base class for all exceptions in the project.
	class Exception
	{
	public:
		Exception(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
		virtual ~Exception()
		{
		}

		virtual String ToString() const;

	protected:
		int    m_nCodeLineNumber;
		String m_sSourceFileName;
		String m_sDebugErrMessage;
	};


	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown by Windows registry helper functions.
	class RegistryException : public Exception
	{
	public:
		RegistryException(LRESULT lErrResult, int nCodeLineNumber = -1,
		                  const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
		virtual String ToString() const;

	private:
		LRESULT m_lRegErrorResult;
	};



	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown in case of invalid parameters.
	class InvalidParamException : public Exception
	{
	public:
		InvalidParamException(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
	};

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown in case the connectin with the browser was lost.
	class BrowserDisconnectedException : public Exception
	{
	public:
		BrowserDisconnectedException(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
	};

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown in case the client cancels the execution.
	class ExecutionCanceledException : public Exception
	{
	public:
		ExecutionCanceledException(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
	};

	///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown in case the operation is not allowed on a certain HTML element (like selectedOption allowed only for combo-boxes).
	class OperationNotAllowedException : public Exception
	{
	public:
		OperationNotAllowedException(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
	};

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exceptions thrown in case the client cancels the execution.
	class IndexOutOfBoundException : public Exception
	{
	public:
		IndexOutOfBoundException(int nCodeLineNumber = -1, const String& sSourceFileName = _T(""), const String& sDebugErrMessage = _T(""));
	};
}


#ifndef DISABLE_TRACE_SERVICE
	#define CreateException(sMessage) \
			ExceptionServices::Exception(__LINE__, __TFILE__, sMessage)

	#define CreateRegistryException(lResult, sMessage) \
		ExceptionServices::RegistryException(lResult, __LINE__, __TFILE__, sMessage)

	#define CreateInvalidParamException(sMessage) \
		ExceptionServices::InvalidParamException(__LINE__, __TFILE__, sMessage)

	#define CreateBrowserDisconnectedException(sMessage) \
		ExceptionServices::BrowserDisconnectedException(__LINE__, __TFILE__, sMessage)

	#define CreateExecutionCanceledException(sMessage) \
		ExceptionServices::ExecutionCanceledException(__LINE__, __TFILE__, sMessage)

	#define CreateExecutionCanceledException(sMessage) \
		ExceptionServices::ExecutionCanceledException(__LINE__, __TFILE__, sMessage)

	#define CreateOperationNotAllowedException(sMessage) \
		ExceptionServices::OperationNotAllowedException(__LINE__, __TFILE__, sMessage)

	#define CreateIndexOutOfBoundException(sMessage) \
		ExceptionServices::IndexOutOfBoundException(__LINE__, __TFILE__, sMessage)
#else
	#define CreateException(sMessage) \
			ExceptionServices::Exception()

	#define CreateRegistryException(lResult, sMessage) \
		ExceptionServices::RegistryException(lResult)

	#define CreateInvalidParamException(sMessage) \
		ExceptionServices::InvalidParamException()

	#define CreateBrowserDisconnectedException(sMessage) \
		ExceptionServices::BrowserDisconnectedException()

	#define CreateExecutionCanceledException(sMessage) \
		ExceptionServices::ExecutionCanceledException()

	#define CreateExecutionCanceledException(sMessage) \
		ExceptionServices::ExecutionCanceledException()

	#define CreateOperationNotAllowedException(sMessage) \
		ExceptionServices::OperationNotAllowedException()

	#define CreateIndexOutOfBoundException(sMessage) \
		ExceptionServices::IndexOutOfBoundException()
#endif //DISABLE_TRACE_SERVICE
