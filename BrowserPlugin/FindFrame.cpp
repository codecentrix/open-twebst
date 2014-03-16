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
#include "HtmlHelpers.h"
#include "FrameVisitor.h"
#include "SearchFrameVisitor.h"
#include "FindFrame.h"

// Local function declarations.
namespace FindFrameAlgorithms
{
	BOOL VisitFrameCollection        (CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor);
	BOOL VisitFrameCollection        (CComQIPtr<IHTMLWindow2> spWindow,  FrameVisitor* pVisitor);
	BOOL VisitChildrenFrameCollection(CComQIPtr<IHTMLWindow2> spWindow,  FrameVisitor* pVisitor);
	BOOL VisitHtmlDialogCollection   (CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor);
	BOOL VisitModalHtmlDialog        (CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor);
	//HWND FindModalHtmlDialog         ();

	BOOL CALLBACK FindAllHtmlDialogWndsCallback(HWND hWnd, LPARAM lParam);
	CComQIPtr<IHTMLWindow2> HTMLWindowFromIeHWND(HWND hIeWnd);
}


list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindFrameAlgorithms::FindAllFrames
                                       (CComQIPtr<IWebBrowser2> spBrowser,
                                        SAFEARRAY* psa, LONG nSearchFlags)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(psa       != NULL);

	SearchFrameCollectionVisitor visitor(psa, nSearchFlags);
	VisitFrameCollection(spBrowser, &visitor);

	return visitor.GetFoundHtmlFrameCollection();
}


list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindFrameAlgorithms::FindAllFrames
                                       (CComQIPtr<IHTMLWindow2> spWindow,
                                        SAFEARRAY* psa, LONG nSearchFlags)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(psa      != NULL);

	SearchFrameCollectionVisitor visitor(psa, nSearchFlags);
	VisitFrameCollection(spWindow, &visitor);

	return visitor.GetFoundHtmlFrameCollection();
}


list<CAdapt<CComQIPtr<IHTMLWindow2> > > FindFrameAlgorithms::FindChildrenFrames
                                       (CComQIPtr<IHTMLWindow2> spWindow,
                                        SAFEARRAY* psa, LONG nSearchFlags)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(psa      != NULL);

	SearchFrameCollectionVisitor visitor(psa, nSearchFlags);
	VisitChildrenFrameCollection(spWindow, &visitor);

	return visitor.GetFoundHtmlFrameCollection();
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::FindFrame(CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(psa       != NULL);

	SearchFrameVisitor visitor(psa);
	VisitFrameCollection(spBrowser, &visitor);

	return visitor.GetFoundHtmlFrame();
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::FindFrame(CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(psa      != NULL);

	SearchFrameVisitor visitor(psa);
	VisitFrameCollection(spWindow, &visitor);

	return visitor.GetFoundHtmlFrame();
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::FindChildFrame(CComQIPtr<IHTMLWindow2> spWindow,  SAFEARRAY* psa)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(psa      != NULL);

	SearchFrameVisitor visitor(psa);
	VisitChildrenFrameCollection(spWindow, &visitor);

	return visitor.GetFoundHtmlFrame();
}



BOOL FindFrameAlgorithms::VisitFrameCollection(CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(pVisitor  != NULL);

	// Get a IHTMLWindow2 object from IWebBrowser2 object.
	CComQIPtr<IHTMLWindow2> spTopFrame = HtmlHelpers::HtmlWebBrowserToHtmlWindow(spBrowser);
	if (spTopFrame != NULL)
	{
		BOOL bRes = pVisitor->VisitFrame(spTopFrame);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}

		return VisitFrameCollection(spTopFrame, pVisitor);
	}
	else
	{
		throw CreateException(_T("Can not get the IHTMLWindow2 from IWebBrowser2 in FindFrameAlgorithms::VisitFrameCollection"));
	}

	return TRUE;
}


BOOL FindFrameAlgorithms::VisitFrameCollection(CComQIPtr<IHTMLWindow2> spWindow,  FrameVisitor* pVisitor)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(pVisitor != NULL);

	BOOL bRes = VisitChildrenFrameCollection(spWindow, pVisitor);
	if (!bRes)
	{
		// Stop browsing.
		return FALSE;
	}

	// Get the frame collection of the spWindow.
	CComQIPtr<IHTMLFramesCollection2> spChildFramesCollection;
	HRESULT hRes = spWindow->get_frames(&spChildFramesCollection);
	if ((hRes != S_OK) || (spChildFramesCollection == NULL))
	{
		throw CreateException(_T("IHTMLWindow2::get_frames failed in FindFrameAlgorithms::VisitFrameCollection"));
	}

	// Get the number of frames in the collection.
	long nCollectionSize;
	hRes = spChildFramesCollection->get_length(&nCollectionSize);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLFramesCollection2::get_length failed in FindFrameAlgorithms::VisitFrameCollection"));
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		// Get the current frame in the collection.
		CComVariant varDispFrame;
		CComVariant varCurrentIndex(i);
		hRes = spChildFramesCollection->item(&varCurrentIndex, &varDispFrame);

		if ((hRes != S_OK) || (varDispFrame.vt != VT_DISPATCH) || (NULL == varDispFrame.pdispVal))
		{
			traceLog << "Can not get the" << i << "th item in FindFrameAlgorithms::VisitFrameCollection\n";
			continue;
		}

		CComQIPtr<IHTMLWindow2> spCurrentFrame = varDispFrame.pdispVal;
		if (spCurrentFrame == NULL)
		{
			traceLog << "Query for IHTMLWindow2 from IDispatch failed for index" << i << " in FindFrameAlgorithms::VisitFrameCollection\n";
			continue;
		}

		// Visit the current html frame.
		bRes = VisitFrameCollection(spCurrentFrame, pVisitor);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


BOOL FindFrameAlgorithms::VisitChildrenFrameCollection(CComQIPtr<IHTMLWindow2> spWindow,  FrameVisitor* pVisitor)
{
	ATLASSERT(spWindow != NULL);
	ATLASSERT(pVisitor != NULL);

	// Get the frame collection of the spWindow.
	CComQIPtr<IHTMLFramesCollection2> spChildFramesCollection;
	HRESULT hRes = spWindow->get_frames(&spChildFramesCollection);
	if ((hRes != S_OK) || (spChildFramesCollection == NULL))
	{
		throw CreateException(_T("IHTMLWindow2::get_frames failed in FindFrameAlgorithms::VisitChildrenFrameCollection"));
	}

	// Get the number of frames in the collection.
	long nCollectionSize;
	hRes = spChildFramesCollection->get_length(&nCollectionSize);
	if (hRes != S_OK)
	{
		throw CreateException(_T("IHTMLFramesCollection2::get_length failed in FindFrameAlgorithms::VisitChildrenFrameCollection"));
	}

	// Browse the collection.
	for (long i = 0; i < nCollectionSize; ++i)
	{
		// Get the current frame in the collection.
		CComVariant varDispFrame;
		CComVariant varCurrentIndex(i);
		hRes = spChildFramesCollection->item(&varCurrentIndex, &varDispFrame);

		if ((hRes != S_OK) || (varDispFrame.vt != VT_DISPATCH) || (NULL == varDispFrame.pdispVal))
		{
			traceLog << "Can not get the" << i << "th item in FindFrameAlgorithms::VisitChildrenFrameCollection\n";
			continue;
		}

		CComQIPtr<IHTMLWindow2> spCurrentFrame = varDispFrame.pdispVal;
		if (spCurrentFrame == NULL)
		{
			traceLog << "Query for IHTMLWindow2 from IDispatch failed for index" << i << " in FindFrameAlgorithms::VisitChildrenFrameCollection\n";
			continue;
		}

		// Visit the current html frame.
		BOOL bRes = pVisitor->VisitFrame(spCurrentFrame);
		if (!bRes)
		{
			// Stop browsing.
			return FALSE;
		}
	}

	return TRUE;
}


struct HtmlDialogInfo
{
public:
	HtmlDialogInfo(HWND hStartSearchWnd, std::list<HWND>* pResultList, BOOL bSearchModeless) :
		m_hStartSearchWnd(hStartSearchWnd), m_pResultList(pResultList), m_bSearchModeless(bSearchModeless)
	{
		ATLASSERT(::IsWindow(hStartSearchWnd));
		ATLASSERT((bSearchModeless && pResultList) || (!bSearchModeless && !pResultList));

		m_hResultModalDlg = NULL;
	}

	BOOL             m_bSearchModeless;
	HWND             m_hResultModalDlg;
	HWND             m_hStartSearchWnd;
	std::list<HWND>* m_pResultList;
};


BOOL CALLBACK FindFrameAlgorithms::FindAllHtmlDialogWndsCallback(HWND hWnd, LPARAM lParam)
{
	ATLASSERT(lParam != NULL);

	if (::IsWindow(hWnd))
	{
		if (Common::GetWndClass(hWnd) == _T("Internet Explorer_TridentDlgFrame"))
		{
			HtmlDialogInfo* pInfo = (HtmlDialogInfo*)lParam;

			// It seems that the parent of modal HTML boxes is no longer IEFrame but a hidden "Alternate Modal Top Most" class.
			// This seems to be the case on Vista + IE8. For modeless dialogs everything remains the same.
			HWND hParentWnd = ::GetParent(hWnd);
			if (!pInfo->m_bSearchModeless || (::IsWindow(hParentWnd) && (hParentWnd == pInfo->m_hStartSearchWnd)))
			{
				HWND hIeWnd = ::FindWindowEx(hWnd, NULL, _T("Internet Explorer_Server"), NULL);
				if (::IsWindow(hIeWnd))
				{
					// Modal HTML dialog boxes are in the same thread as the tab. Not the case for modeless dialogs.
					DWORD dwIEWndThread = ::GetWindowThreadProcessId(hIeWnd, NULL);
					DWORD dwCrntTrhread = ::GetCurrentThreadId();

					if (pInfo->m_bSearchModeless)
					{
						// Add only modeless dialogs to the list.
						if (dwIEWndThread != dwCrntTrhread)
						{
							std::list<HWND>* pList = pInfo->m_pResultList;
							pList->push_back(hIeWnd);
						}
					}
					else
					{
						if (dwIEWndThread == dwCrntTrhread)
						{
							pInfo->m_hResultModalDlg = hIeWnd;
							return FALSE;
						}
					}
				}
			}
		}
	}

	return TRUE;
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::HTMLWindowFromIeHWND(HWND hIeWnd)
{
	ATLASSERT(::IsWindow(hIeWnd));
	ATLASSERT(Common::GetWndClass(hIeWnd) == _T("Internet Explorer_Server"));

	// Get the accessible object from the "Internet Explorer_Server" window.
	CComQIPtr<IAccessible> spAccObj;
	HRESULT hRes = ::AccessibleObjectFromWindow(hIeWnd, OBJID_WINDOW, IID_IAccessible, (void**)&spAccObj);

	if (SUCCEEDED(hRes) && (spAccObj != NULL))
	{
		// Get the service provider from the accessible object.
		CComQIPtr<IServiceProvider>	spServProv = spAccObj;

		if (spServProv != NULL)
		{
			// Query the service provider for IHTMLWindow2 interface
			CComQIPtr<IHTMLWindow2> spHtmlWindow;
			hRes = spServProv->QueryService(IID_IHTMLWindow2, IID_IHTMLWindow2, (void**)&spHtmlWindow);

			if (SUCCEEDED(hRes) && (spHtmlWindow != NULL))
			{
				return spHtmlWindow;
			}
			else
			{
				traceLog << "Can NOT get IHTMLWindow2 in FindFrameAlgorithms::HTMLWindowFromIeHWND\n";
			}
		}
		else
		{
			traceLog << "Can NOT get IServiceProvider in FindFrameAlgorithms::HTMLWindowFromIeHWND\n";
		}
	}
	else
	{
		traceLog << "AccessibleObjectFromWindow failed in FindFrameAlgorithms::HTMLWindowFromIeHWND code:" << hRes << "\n";
	}

	return CComQIPtr<IHTMLWindow2>();
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::FindHtmlDialog(CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(psa       != NULL);

	SearchFrameVisitor visitor(psa);
	VisitHtmlDialogCollection(spBrowser, &visitor);

	return visitor.GetFoundHtmlFrame();
}


CComQIPtr<IHTMLWindow2> FindFrameAlgorithms::FindModalHtmlDialog(CComQIPtr<IWebBrowser2> spBrowser, SAFEARRAY* psa)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(psa       != NULL);

	SearchFrameVisitor visitor(psa);
	VisitModalHtmlDialog(spBrowser, &visitor);

	return visitor.GetFoundHtmlFrame();
}


BOOL FindFrameAlgorithms::VisitHtmlDialogCollection(CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(pVisitor  != NULL);

	HWND hStartSearchWnd = NULL;
	BOOL bErr            = FALSE;

	// Use the main IEFrame window as starting point for enumerating IE server children windows.
	BOOL     bInsideIE = TRUE;
	LONG_PTR nTopWnd   = NULL;
	HRESULT  hRes      = spBrowser->get_HWND(&nTopWnd);

	if (SUCCEEDED(hRes))
	{
		hStartSearchWnd = (HWND)nTopWnd;
	}
	else
	{
		bInsideIE = FALSE;
		traceLog << "get_HWND failed in CBrowser::VisitHtmlDialogCollection\n";

		// It can fail for embeded browser controls.
		HWND hTabWnd = HtmlHelpers::GetIEWndFromBrowser(spBrowser);
		hStartSearchWnd = Common::GetTopParentWnd(hTabWnd);

		if (!::IsWindow(hStartSearchWnd))
		{
			bErr = TRUE;
			traceLog << "GetIEWndFromBrowser, Common::GetTopParentWnd in CBrowser::VisitHtmlDialogCollection\n";
		}
	}

	if (bErr || !::IsWindow(hStartSearchWnd))
	{
		throw CreateException(_T("Can not get the hwnd to start searching in FindFrameAlgorithms::VisitHtmlDialogCollection"));
	}

	String sClassName = Common::GetWndClass(hStartSearchWnd);
	if (bInsideIE && (sClassName != _T("IEFrame")))
	{
		throw CreateException(_T("Invalid start hwnd found in FindFrameAlgorithms::VisitHtmlDialogCollection"));
	}

	// Enumerate all "Internet Explorer_TridentDlgFrame" children windows.
	std::list<HWND> ieWndList;
	HtmlDialogInfo  htmlDialogInfo(hStartSearchWnd, &ieWndList, TRUE);

	::EnumWindows(FindAllHtmlDialogWndsCallback, (LPARAM)(&htmlDialogInfo));

	for (std::list<HWND>::iterator it = ieWndList.begin();
		 it != ieWndList.end(); ++it)
	{
		HWND hIEWnd = *it;

		ATLASSERT(Common::GetWndClass(hIEWnd) == _T("Internet Explorer_Server"));

		CComQIPtr<IHTMLWindow2> spCrntHtmlWnd = FindFrameAlgorithms::HTMLWindowFromIeHWND(hIEWnd);
		if (spCrntHtmlWnd != NULL)
		{
			BOOL bRes = pVisitor->VisitFrame(spCrntHtmlWnd);

			if (!bRes)
			{
				// Stop browsing.
				return FALSE;
			}
		}
	}

	return TRUE;
}


// This was just an alternative to EnumWindows.
/*HWND FindFrameAlgorithms::FindModalHtmlDialog()
{
	HWND hCrntChildWnd = NULL;
	HWND hModalWnd     = NULL;

	while (TRUE)
	{
		hModalWnd = ::FindWindowEx(NULL, hCrntChildWnd, _T("Internet Explorer_TridentDlgFrame"), NULL);
		if (NULL == hModalWnd)
		{
			break;
		}

		hCrntChildWnd = hModalWnd;

		// Modal HTML dialog boxes are in the same thread as the tab. Not the case for modeless dialogs.
		DWORD dwIEWndThread = ::GetWindowThreadProcessId(hModalWnd, NULL);
		DWORD dwCrntTrhread = ::GetCurrentThreadId();

		if (dwIEWndThread == dwCrntTrhread)
		{
			break;
		}
	}

	return hModalWnd;
}*/


BOOL FindFrameAlgorithms::VisitModalHtmlDialog(CComQIPtr<IWebBrowser2> spBrowser, FrameVisitor* pVisitor)
{
	ATLASSERT(spBrowser != NULL);
	ATLASSERT(pVisitor  != NULL);

	HWND hStartSearchWnd = NULL;
	BOOL bErr            = FALSE;

	// Use the main IEFrame window as starting point for enumerating IE server children windows.
	BOOL     bInsideIE = TRUE;
	LONG_PTR nTopWnd   = NULL;
	HRESULT  hRes      = spBrowser->get_HWND(&nTopWnd);

	if (SUCCEEDED(hRes))
	{
		hStartSearchWnd = (HWND)nTopWnd;
	}
	else
	{
		bInsideIE = FALSE;
		traceLog << "get_HWND failed in CBrowser::VisitModalHtmlDialog\n";

		// It can fail for embeded browser controls.
		HWND hTabWnd = HtmlHelpers::GetIEWndFromBrowser(spBrowser);
		hStartSearchWnd = Common::GetTopParentWnd(hTabWnd);

		if (!::IsWindow(hStartSearchWnd))
		{
			bErr = TRUE;
			traceLog << "GetIEWndFromBrowser, Common::GetTopParentWnd in CBrowser::VisitModalHtmlDialog\n";
		}
	}

	if (bErr || !::IsWindow(hStartSearchWnd))
	{
		throw CreateException(_T("Can not get the hwnd to start searching in FindFrameAlgorithms::VisitModalHtmlDialog"));
	}

	String sClassName = Common::GetWndClass(hStartSearchWnd);
	if (bInsideIE && (sClassName != _T("IEFrame")))
	{
		throw CreateException(_T("Invalid start hwnd found in FindFrameAlgorithms::VisitModalHtmlDialog"));
	}

	// Just alternative try that worked.
	//HWND hIEWnd = FindModalHtmlDialog();

	HtmlDialogInfo htmlDialogInfo(hStartSearchWnd, NULL, FALSE);
	::EnumWindows(FindAllHtmlDialogWndsCallback, (LPARAM)(&htmlDialogInfo));

	HWND hIEWnd = htmlDialogInfo.m_hResultModalDlg;
	if (hIEWnd != NULL)
	{
		ATLASSERT(Common::GetWndClass(hIEWnd) == _T("Internet Explorer_Server"));

		CComQIPtr<IHTMLWindow2> spCrntHtmlWnd = FindFrameAlgorithms::HTMLWindowFromIeHWND(hIEWnd);
		if (spCrntHtmlWnd != NULL)
		{
			BOOL bRes = pVisitor->VisitFrame(spCrntHtmlWnd);

			if (!bRes)
			{
				// Stop browsing.
				return FALSE;
			}
		}
	}

	return TRUE;
}
