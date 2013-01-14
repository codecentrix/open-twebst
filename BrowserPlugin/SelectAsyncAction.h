#pragma once
#include "AsyncAction.h"

class CExplorerPlugin;

class SelectAsyncAction : public IAsyncAction
{
public:
	SelectAsyncAction(CExplorerPlugin* pPlugin, CComQIPtr<IHTMLElement> spTargetElement, const list<DWORD>& itemsToSelect, BOOL bIsCombo, BOOL bMultiple, LONG nFlags);
	virtual HRESULT DoAction ();
	virtual ~SelectAsyncAction();

private:
	CExplorerPlugin*        m_pPlugin;
	CComQIPtr<IHTMLElement> m_spTargetElement;
	list<DWORD>             m_itemsToSelect;

	BOOL m_bIsCombo;
	BOOL m_bMultiple;
	LONG m_nFlags;
};
