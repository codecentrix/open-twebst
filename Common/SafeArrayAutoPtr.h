#pragma once

class SafeArrrayAutoPtr
{
public:
	SafeArrrayAutoPtr() : m_pSafeArray(NULL) {}

	void Attach(SAFEARRAY* pSafeArray)
	{
		ATLASSERT(pSafeArray != NULL);
		m_pSafeArray = pSafeArray;
	}

	~SafeArrrayAutoPtr()
	{
		if (m_pSafeArray != NULL)
		{
			HRESULT hRes = ::SafeArrayDestroy(m_pSafeArray);
			ATLASSERT(SUCCEEDED(hRes));
		}
	}

private:
	SAFEARRAY* m_pSafeArray;
};
