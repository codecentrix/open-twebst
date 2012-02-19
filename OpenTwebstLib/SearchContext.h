#pragma once

struct SearchContext
{
public:
	SearchContext(const std::string& sFunctionName, ULONG nSearchContextFlags, DWORD dwHelpContextID,
                  DWORD dwInvalidParamMsgID, DWORD dwFailureParamMsgID, DWORD dwTimeoutMsgID) :
		m_sFunctionName  (sFunctionName),      m_nSearchContextFlags(nSearchContextFlags),
		m_dwHelpContextID(dwHelpContextID),    m_dwInvalidParamMsgID(dwInvalidParamMsgID),
		m_dwFailureMsgID(dwFailureParamMsgID), m_dwTimeoutMsgID(dwTimeoutMsgID)
	{
	}

	std::string m_sFunctionName;
	ULONG  m_nSearchContextFlags;
	DWORD  m_dwHelpContextID;
	DWORD  m_dwInvalidParamMsgID;
	DWORD  m_dwFailureMsgID;
	DWORD  m_dwTimeoutMsgID;
};
