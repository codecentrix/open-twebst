///////////////////////////////////////////////////////////////////////////////
// (c)2012 CodeCentrix Software. All rights reserved.
//
// This sample demonstrates:
// - Creating the Core object that represents the root of the Open Twebst library.
// - Starting a new Browser and navigate to a given url.
// - Finding HTML elements inside the browser.
// - Performing actions on HTML controls.
///////////////////////////////////////////////////////////////////////////////

using System;
using mshtml;
using Core     = OpenTwebstLib.core;
using ICore    = OpenTwebstLib.ICore;
using IBrowser = OpenTwebstLib.IBrowser;
using IElement = OpenTwebstLib.IElement;


namespace OpenTwebstSample
{
	public class Sample
	{
		// The main entry point for the application.
		[STAThread]
		static void Main() 
		{
			// Create the Core object.
			ICore core = new Core();

            // Use hardware events to perform actions on HTML controls because of coolness factor.
            // By default HTML events are used.
            core.useHardwareInputEvents = true;

            // Start a new Internet Explorer browser and create a Browser object for it.
            IBrowser browser = core.StartBrowser("http://doc.codecentrix.com/Lnkplayground.htm");

            System.Threading.Thread.Sleep(1000);

            // Find "Email" editbox control.
            IElement elem = browser.FindElement("input text", "id=email");
            elem.InputText("open-twebst@codecentrix.com");

            // Find "Password" editbox control.
            elem = browser.FindElement("input text", "id=password");
            elem.InputText("notarealpwd");

            // Find "Login" button control.
            elem = browser.FindElement("input submit", "id=login");
            elem.Click();

            // Close popup window.
            System.Threading.Thread.Sleep(1000);
            browser.ClosePopup("", 0);

            // Automate a list box.
            elem = browser.FindElement("select", "id=cntry");
            elem.Select("Romania");

            // Automate check box and radio.
            elem = browser.FindElement("input radio", "id=r3");
            elem.Check();

            elem = browser.FindElement("input checkbox", "");
            elem.Check();

            // Modify the HTML document.
            elem = browser.FindElement("div", "id=deev");
            elem.nativeElement.innerText = "just changed";

            // Automate inside iframe.
            elem = browser.FindElement("a", "uiName=Downloads");
            elem.Click();
		}
	}
}
