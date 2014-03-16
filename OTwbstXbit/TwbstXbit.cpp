// TwbstXbit.cpp : Defines the entry point for the application.
#include "stdafx.h"
#include "TwbstXbit.h"
#include "Common.h"
#include "..\OTWBSTInjector\OTWBSTInjector.h"
#include "MarshalService.h"

using namespace XBit;

// Shared variable.
#ifdef _AMD64_
	#pragma data_seg("OTwbstXbitShMemx64")
		LONG g_nInstNo = -1;
	#pragma data_seg()

	#pragma comment(linker, "/SECTION:OTwbstXbitShMemx64,RWS")
#else
	#pragma data_seg("OTwbstXbitShMemx86")
		LONG g_nInstNo = -1;
	#pragma data_seg()

	#pragma comment(linker, "/SECTION:OTwbstXbitShMemx86,RWS")
#endif


class CInjectorWnd : public CWindowImpl<CInjectorWnd>
{
public:
	CInjectorWnd()
	{
		m_lastRequest = ::GetTickCount();
	}

private:
	BEGIN_MSG_MAP(CInjectorWnd)
		MESSAGE_HANDLER(XBIT_MSG_INJECT_BHO , OnInjectBHO)
	END_MSG_MAP()

#ifdef _AMD_64_
	DECLARE_WND_CLASS(XBit::XBIT_INJECTOR_WND_CLASS_NAME_x64);
#else
	DECLARE_WND_CLASS(XBit::XBIT_INJECTOR_WND_CLASS_NAME_x86);
#endif

	LRESULT OnInjectBHO(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

private:
	DWORD m_lastRequest;
};


LRESULT CInjectorWnd::OnInjectBHO(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	m_lastRequest = ::GetTickCount();
	bHandled = TRUE;

	BOOL bRes = InjectBHO((HWND)lParam, static_cast<BOOL>(wParam), CComQIPtr<IWebBrowser2>());
	return bRes;
}


int StartIE(LPCWSTR szUrl)
{
	if (!szUrl)
	{
		return 0;
	}

	return Common::StartInternetExplorerProcessOnVista(szUrl, NULL);
}



//////////////////////////////////////////////////////////////////////////////////////////////////////////
int APIENTRY _tWinMain
	(
		HINSTANCE hInstance,
        HINSTANCE hPrevInstance,
        LPTSTR    lpCmdLine,
        int       nCmdShow
	)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// Process command line.
	int     nArgs     = 0;
	LPWSTR* szArglist = ::CommandLineToArgvW(GetCommandLineW(), &nArgs);

	if (szArglist != NULL)
	{
		if ((nArgs > 1) && !_wcsicmp(XBIT_START_IE_CMDL, szArglist[0]))
		{
			return StartIE(szArglist[1]);
		}

		// Clean up.
		LocalFree(szArglist);
		szArglist = NULL;
	}

	BOOL bFirstInst = (InterlockedIncrement(&g_nInstNo) == 0);
	if (!bFirstInst)
	{
		InterlockedDecrement(&g_nInstNo);
		return 0;
	}

	// Create communication window.
	CInjectorWnd injectWnd;
	injectWnd.Create(NULL, NULL, XBit::XBIT_INJECTOR_WND_NAME, WS_POPUP);

	MSG msg = {};

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	injectWnd.DestroyWindow();
	return -1;
}
