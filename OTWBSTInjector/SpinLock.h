#pragma once

const int SL_LOCKED   = 1;
const int SL_UNLOCKED = 0;


class SpinLock
{
public:
	explicit SpinLock(volatile long* pLock)
	{
		_ASSERTE(pLock != NULL);
		m_pLock = pLock;
	}

	~SpinLock()
	{
		m_pLock = NULL;
	}

	void Lock()
    {
        while (::InterlockedExchange(m_pLock, SL_LOCKED) != SL_UNLOCKED)
        {
			::Sleep(30);
        }
	}

	void Unlock()
    {
		_ASSERTE(HasLock());
		::InterlockedExchange(m_pLock, SL_UNLOCKED);
	}


private:
    SpinLock(const SpinLock &);
    SpinLock& operator=(const SpinLock &);

	BOOL HasLock() const
	{
		return (*m_pLock == SL_LOCKED);
	}


private:
	volatile long* m_pLock;
};
