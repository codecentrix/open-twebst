///////////////////////////////////////////////////////////////////////////////
// (c)2012 CodeCentrix Software. All rights reserved.
//
// This sample demonstrates:
// - Creating the Core object that represents the root of the Open Twebst library.
// - Starting a new Browser and navigate to a given url.
// - Finding HTML elements inside the browser.
// - Performing actions on HTML controls.
///////////////////////////////////////////////////////////////////////////////


// Instantiate a Core object.
var core = new ActiveXObject("OpenTwebst.Core");

// Use hardware events to perform actions on HTML controls because of coolness factor.
// By default HTML events are used.
core.useHardwareInputEvents = true;

// Start a new Internet Explorer browser and create a Browser object for it.
var browser = core.StartBrowser("http://doc.codecentrix.com/playground.htm");

WScript.Sleep(1000);

// Find "Email" editbox control.
var elem = browser.FindElement("input text", "id=email");
elem.Highlight();
elem.InputText("open-twebst@codecentrix.com");

// Find "Password" editbox control.
elem = browser.FindElement("input text", "id=password");
elem.Highlight();
elem.InputText("notarealpwd");

// Find "Login" button control.
elem = browser.FindElement("input submit", "id=login");
elem.Highlight();
elem.Click();

// Close popup window.
WScript.Sleep(1000);
browser.ClosePopup("", 0);

// Automate a list box.
elem = browser.FindElement("select", "id=cntry");
elem.Highlight();
elem.Select("Romania");

// Automate check box and radio.
elem = browser.FindElement("input radio", "id=r3");
elem.Highlight();
elem.Check();

elem = browser.FindElement("input checkbox", "");
elem.Highlight();
elem.Check();

// Modify the HTML document.
elem = browser.FindElement("div", "id=deev");
elem.Highlight();
elem.nativeElement.innerText = "Changed by web automation script!";

// Automate inside iframe.
elem = browser.FindElement("a", "uiName=Downloads");
elem.Highlight();
elem.Click();
