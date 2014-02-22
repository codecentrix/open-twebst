using System;
using System.Windows.Forms;
using OpenTwebstLib;



namespace TwebstWebBrowser
{
    public partial class TwebstForm : Form
    {
        public TwebstForm()
        {
            InitializeComponent();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            // Twebst automation code for hosted WebBrowser control.

            // Attach to native IWebBrowser2 object. If this code runs in other thread than UI thread you need
            // to marshal IWebBrowser2 object; see CoMarshalInterThreadInterfaceInStream in MSDN
            // http://msdn.microsoft.com/en-us/library/windows/desktop/ms693316(v=vs.85).aspx
            IBrowser browser = core.AttachToNativeBrowser((SHDocVw.IWebBrowser2)this.webBrowser.ActiveXInstance);

            // Or you can attach directly to Win32 window handle. Reccomended if you automate WebBrowser
            // from another thread because you don't need to marshal any COM object.
            // IBrowser browser = core.AttachToHWND(this.webBrowser.Handle.ToInt32());

            browser.Navigate("http://codecentrix.com/doc/playground.htm");

            // Find "Email" editbox control.
            IElement elem = browser.FindElement("input text", "id=email");
            elem.InputText("open-twebst@codecentrix.com");

            // Find "Password" editbox control.
            elem = browser.FindElement("input text", "id=password");
            elem.InputText("notarealpwd");

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
            elem = browser.FindElement("a", "uiName=Download");
            elem.Click();

            // !!!!!! ALERT:
            // Some features don't work with hosted WebBrowser yet: closePopup, upload controls, modal and modeless HTML dialog boxes.
        }

        private void OnCancel(out bool cancel)
        {
            // The library fires CancelRequest COM events while waiting for the page to be completely loaded or while searching for HTML objects.
            // You can cancel Twebst automation code by setting cancel to true; the current Twebst method will throw an exception.
            // You do not need to call Application.DoEvents() like in earlier versions; Twebst methods process windows messages by themselves.

            cancelReqCount++;
            System.Diagnostics.Trace.WriteLine("Cancel request, count=" + cancelReqCount);

            cancel = false;
        }

        private void TwebstForm_Load(object sender, EventArgs e)
        {
            this.webBrowser.Navigate("about:blank");

            // Register to cancel event.
            ((OpenTwebstLib.core)this.core).CancelRequest += OnCancel;
        }

        private ICore core = new core();
        private int cancelReqCount = 0;
    }
}
