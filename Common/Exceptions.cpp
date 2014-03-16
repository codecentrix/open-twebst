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


namespace ExceptionServices
{
	////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exception
	Exception::Exception(int nCodeLineNumber, const String& sSourceFileName, const String& sDebugErrMessage) :
		m_nCodeLineNumber(nCodeLineNumber),
		m_sSourceFileName(sSourceFileName),
		m_sDebugErrMessage(sDebugErrMessage)
	{
		
	}

	String Exception::ToString() const
	{
		Common::Ostringstream outputStream;
		
		outputStream << _T("'") << m_sDebugErrMessage << _T("' ") << m_sSourceFileName  << _T("(") << m_nCodeLineNumber << _T(")");
		return outputStream.str();
	}


	////////////////////////////////////////////////////////////////////////////////////////////////////
	// RegistryException
	RegistryException::RegistryException(LRESULT lErrResult, int nCodeLineNumber,
	                                     const String& sSourceFileName, const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage), m_lRegErrorResult(lErrResult)
	{
		
	}

	String RegistryException::ToString() const
	{
		Common::Ostringstream outputStream;

		outputStream << Exception::ToString() << _T("lResult= ") << m_lRegErrorResult;
		return outputStream.str();
	}


	////////////////////////////////////////////////////////////////////////////////////////////////////
	// InvalidParamException
	InvalidParamException::InvalidParamException(int   nCodeLineNumber,
                                                 const String& sSourceFileName,
                                                 const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage)
	{
		
	}


	////////////////////////////////////////////////////////////////////////////////////////////////////
	// BrowserDisconnectedException
	BrowserDisconnectedException::BrowserDisconnectedException(int nCodeLineNumber,
                                                 const String& sSourceFileName,
                                                 const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage)
	{
		
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// ExecutionCanceledException
	ExecutionCanceledException::ExecutionCanceledException(int nCodeLineNumber,
                                                 const String& sSourceFileName,
                                                 const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage)
	{
		
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// OperationNotAllowedException
	OperationNotAllowedException::OperationNotAllowedException(int nCodeLineNumber,
                                                 const String& sSourceFileName,
                                                 const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage)
	{
		
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// IndexOutOfBoundException
	IndexOutOfBoundException::IndexOutOfBoundException(int nCodeLineNumber,
                                                 const String& sSourceFileName,
                                                 const String& sDebugErrMessage) :
			Exception(nCodeLineNumber, sSourceFileName, sDebugErrMessage)
	{
		
	}
}
