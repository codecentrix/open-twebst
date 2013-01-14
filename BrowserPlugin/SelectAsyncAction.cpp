#include "StdAfx.h"
#include "SelectAsyncAction.h"
#include "ExplorerPlugin.h"



SelectAsyncAction::SelectAsyncAction
	(
		CExplorerPlugin*        pPlugin,
		CComQIPtr<IHTMLElement> spTargetElement,
		const list<DWORD>&      itemsToSelect,
		BOOL                    bIsCombo,
		BOOL                    bMultiple,
		LONG                    nFlags
	) :
	m_pPlugin(pPlugin), m_spTargetElement(spTargetElement), m_itemsToSelect(itemsToSelect), m_bIsCombo(bIsCombo), m_bMultiple(bMultiple), m_nFlags(nFlags)
{
	ATLASSERT(pPlugin         != NULL);
	ATLASSERT(spTargetElement != NULL);
	ATLASSERT(itemsToSelect.size() > 0);
}


HRESULT SelectAsyncAction::DoAction()
{
	ATLASSERT(m_pPlugin != NULL);
	return m_pPlugin->DoSelectOptions(m_spTargetElement, m_itemsToSelect, m_bIsCombo, m_bMultiple, m_nFlags);
}


SelectAsyncAction::~SelectAsyncAction()
{
}
