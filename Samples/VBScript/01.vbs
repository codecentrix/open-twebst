'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (c)2012 CodeCentrix Software. All rights reserved.
'
' This sample demonstrates:
' - Creating the Core object that represents the root of the Twebst library.
' - Starting a new Browser and navigate to a given url.
' - Finding HTML elements inside the browser.
' - Performing actions on HTML controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim core
Dim browser
Dim elem

' Instantiate a Twebst Core object.
Set core = CreateObject("OpenTwebst.Core")

' Use hardware events to perform actions on HTML controls because of coolness factor.
' By default HTML events are used.
On Error Resume Next
core.useHardwareInputEvents = true

' Start a new Internet Explorer browser and create a Twebst Browser object for it.
Set browser = core.StartBrowser("http://doc.codecentrix.com/Lnkplayground.htm")

Call WScript.Sleep(1000)

' Find "Email" editbox control.
Set elem = browser.FindElement("input text", "id=email")
Call elem.InputText("open-twebst@codecentrix.com")

' Find "Password" editbox control.
Set elem = browser.FindElement("input text", "id=password")
Call elem.InputText("notarealpwd")

' Find "Login" button control.
Set elem = browser.FindElement("input submit", "id=login")
Call elem.Click()

' Close popup window.
Call WScript.Sleep(1000)
Call browser.ClosePopup("", 0)

' Automate a list box.
Set elem = browser.FindElement("select", "id=cntry")
Call elem.Select("Romania")

' Automate check box and radio.
Set elem = browser.FindElement("input radio", "id=r3")
Call elem.Check()

Set elem = browser.FindElement("input checkbox", "")
Call elem.Check()

' Modify the HTML document.
Set elem = browser.FindElement("div", "id=deev")
Set elem.nativeElement.innerText = "just changed"

' Automate inside iframe.
Set elem = browser.FindElement("a", "uiName=Downloads")
Call elem.Click()
