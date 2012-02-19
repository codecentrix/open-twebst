////////////////////////////////////////////////////////////////
// (c)2012 CodeCentrix Software. All rights reserved.
//
// This C++ application demonstrates the main features of
// Open Twebst Library like form filling and web automation.
////////////////////////////////////////////////////////////////

#include "stdafx.h"
#import "progid:OpenTwebst.Core"

using namespace std;
using namespace OpenTwebstLib;


int _tmain(int argc, _TCHAR* argv[])
{
	::CoInitialize(NULL);

	try
	{
		// Create the Core object.
		ICorePtr pCore;
		HRESULT  hRes = pCore.CreateInstance(_T("OpenTwebst.Core"));
		if (FAILED(hRes))
		{
			// Failed to create the Core object.
			throw _com_error(hRes);
		}

		// Use hardware events to perform actions on HTML controls because of coolness factor.
		// By default HTML events are used.
		pCore->useHardwareInputEvents = VARIANT_TRUE;

		// Start a new Internet Explorer instance and navigate to a given URL.
		IBrowserPtr pBrowser = pCore->StartBrowser("http://doc.codecentrix.com/Lnkplayground.htm");

		::Sleep(1000);

		// Find "Email" editbox control.
		IElementPtr pElem = pBrowser->FindElement("input text", "id=email");
		pElem->InputText("open-twebst@codecentrix.com");

		// Find "Password" editbox control.
		pElem = pBrowser->FindElement("input text", "id=password");
		pElem->InputText("notarealpwd");

		// Find "Login" button control.
		pElem = pBrowser->FindElement("input submit", "id=login");
		pElem->Click();

		// Close popup window.
		::Sleep(1000);
		pBrowser->ClosePopup("", 0);

		// Automate a list box.
		pElem = pBrowser->FindElement("select", "id=cntry");
		pElem->Select("Romania");

		// Automate check box and radio.
		pElem = pBrowser->FindElement("input radio", "id=r3");
		pElem->Check();

		pElem = pBrowser->FindElement("input checkbox", "");
		pElem->Check();

		// Modify the HTML document.
		pElem = pBrowser->FindElement("div", "id=deev");
		pElem->nativeElement->put_innerText(CComBSTR("just changed"));

		// Automate inside iframe.
		pElem = pBrowser->FindElement("a", "uiName=Downloads");
		pElem->Click();
	}
	catch (_com_error comErr)
	{
		cout << "Error, hRes=" << std::hex << comErr.Error() << "\n";
	}

	::CoUninitialize();
	return 0;
}
