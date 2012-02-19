using System;
using System.Collections.Generic;
using System.Text;
using mshtml;
using SHDocVw;
using OpenTwebstLib;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Accessibility;
using System.Drawing;


namespace CatStudio
{
    class Recorder : Form
    {
        #region Public section
        public event EventHandler<EventArgs>    RecStarted;
        public event EventHandler<EventArgs>    RecStopped;
        public event EventHandler<RecEventArgs> ClickAction;
        public event EventHandler<RecEventArgs> RightClickAction;
        public event EventHandler<RecEventArgs> ChangeAction;
        public event EventHandler<EventArgs>    BackAction;
        public event EventHandler<EventArgs>    ForwardAction;

        public static Recorder Instance
        {
            get
            {
                return singleInstance;
            }
        }


        public void StartRec()
        {
            this.isRecording = true;

            if (RecStarted != null)
            {
                RecStarted(this, EventArgs.Empty);
            }
        }


        public void StopRec()
        {
            this.isRecording = false;

            if (RecStopped != null)
            {
                RecStopped(this, EventArgs.Empty);
            }
        }


        public bool IsRecording
        {
            get { return this.isRecording; }
        }


        public void RecordForward()
        {
            if (this.isRecording && (this.ForwardAction != null))
            {
                this.ForwardAction(this, EventArgs.Empty);
            }
        }


        public void RecordBack()
        {
            if (this.isRecording && (this.BackAction != null))
            {
                this.BackAction(this, EventArgs.Empty);
            }
        }


        public void Init(System.Windows.Forms.WebBrowser webBrowser)
        {
            this.twebstCore.loadTimeout   = 0;
            this.twebstCore.searchTimeout = 0;

            IWebBrowser2 nativeBrowser = (SHDocVw.IWebBrowser2)webBrowser.ActiveXInstance;
            this.twebstBrowser = this.twebstCore.AttachToNativeBrowser(nativeBrowser);

            this.win32Hook.Win32HookMsg += OnWin32HookMsg;
            this.win32Hook.Install();
        }


        public void Uninit()
        {
            this.win32Hook.Win32HookMsg -= OnWin32HookMsg;
            this.win32Hook.UnInstall();

            isRecording        = false;
            this.twebstBrowser = null;
            this.twebstCore.Reset();
        }
        #endregion

        #region Private section
        private void OnWin32HookMsg(Object sender, Win32HookMsgEventArgs e)
        {
            if (this.isRecording)
            {
                this.pendingWin32Action = true;
                this.lastWin32Time      = DateTime.Now;

                if (e.Type == Win32HookMsgEventArgs.HookMsgType.CLICK_DOWN_MSG)
                {
                    this.pendingWin32Click = true;
                }
                else if (e.Type == Win32HookMsgEventArgs.HookMsgType.CLICK_UP_MSG)
                {
                    //System.Diagnostics.Trace.WriteLine("WM_PENDING_WIN32_CLICK posted because of " + e.Type);

                    Point screenPos = System.Windows.Forms.Cursor.Position;
                    Win32Api.PostMessage(this.Handle, WM_PENDING_WIN32_CLICK, Win32Api.MakeLong((short)screenPos.X, (short)screenPos.Y), e.WndHandle.ToInt32());
                }
                else if (e.Type == Win32HookMsgEventArgs.HookMsgType.RIGHT_CLICK_DOWN_MSG)
                {
                    Point screenPos = System.Windows.Forms.Cursor.Position;
                    OnRightClick(e.WndHandle, screenPos);
                }

                // Install HTML hooks on left click down and key down.
                if ((e.Type != Win32HookMsgEventArgs.HookMsgType.CLICK_UP_MSG)       &&
                    (e.Type != Win32HookMsgEventArgs.HookMsgType.RIGHT_CLICK_UP_MSG) &&
                    (e.Type != Win32HookMsgEventArgs.HookMsgType.RIGHT_CLICK_DOWN_MSG))
                {
                    InstallHTMLHooks(e);
                }
            }
        }


        private void InstallHTMLHooks(Win32HookMsgEventArgs evt)
        {
            String       searchCondition   = String.Format("{0}!=1", CatStudioConstants.HOOKED_BY_REC_ATTR);
            IElementList rootElemNotHooked = this.twebstBrowser.FindAllElements("html", searchCondition);

            try
            {
                for (int i = 0; i < rootElemNotHooked.length; ++i)
                {

                    IElement       crntElem = rootElemNotHooked[i];
                    IHTMLDocument3 doc3     = (IHTMLDocument3)crntElem.nativeElement.document;
                    IHTMLWindow2   wnd      = ((IHTMLDocument2)doc3).parentWindow;

                    doc3.attachEvent("onclick",   new HtmlHandler(this.OnHtmlClick, wnd));
                    doc3.attachEvent("onmouseup", new HtmlHandler(this.OnHtmlMouseUp, wnd));
                    crntElem.SetAttribute(CatStudioConstants.HOOKED_BY_REC_ATTR, "1");
                }

                InstallHTMLHooksForOnchange();
            }
            catch
            {
                // Can not properly install html hooks.
            }
        }


        private void InstallHTMLHooksForOnchange()
        {
            // Install onchange hook for <input type=password, text, file>.
            String       searchCondition    = String.Format("{0}!=1", CatStudioConstants.HOOKED_BY_REC_ATTR);
            IElementList inputElemNotHooked = this.twebstBrowser.FindAllElements("input", searchCondition);

            for (int i = 0; i < inputElemNotHooked.length; ++i)
            {
                IHTMLElement2 crntNativeElem2 = (IHTMLElement2)inputElemNotHooked[i].nativeElement;
                String        inputType       = ((IHTMLInputElement)crntNativeElem2).type.ToLower();
                IHTMLElement  crntElem        = (IHTMLElement)crntNativeElem2;

                if (("text" == inputType) || ("password" == inputType) || ("file" == inputType))
                {
                    IHTMLWindow2 wnd = ((IHTMLDocument2)crntElem.document).parentWindow;
                    crntNativeElem2.attachEvent("onchange", new HtmlHandler(this.OnHtmlChange, wnd));
                }

                crntElem.setAttribute(CatStudioConstants.HOOKED_BY_REC_ATTR, "1", 0);
            }

            // Install onchange hook for <select>.
            IElementList selectElemNotHooked = this.twebstBrowser.FindAllElements("select", searchCondition);
            InstalHTMLHookOnList(selectElemNotHooked, "onchange");

            // Install onchange hook for <textarea>.
            IElementList textAreaElemNotHooked = this.twebstBrowser.FindAllElements("textarea", searchCondition);
            InstalHTMLHookOnList(textAreaElemNotHooked, "onchange");
        }


        private void InstalHTMLHookOnList(IElementList elemList, String evtName)
        {
            for (int i = 0; i < elemList.length; ++i)
            {
                IHTMLElement2 crntNativeElem2 = (IHTMLElement2)elemList[i].nativeElement;
                IHTMLElement  crntElem        = (IHTMLElement)crntNativeElem2;
                IHTMLWindow2  wnd             = ((IHTMLDocument2)crntElem.document).parentWindow;

                crntNativeElem2.attachEvent(evtName, new HtmlHandler(this.OnHtmlChange, wnd));
                crntElem.setAttribute(CatStudioConstants.HOOKED_BY_REC_ATTR, "1", 0);
            }
        }


        private void OnHtmlChange(object sender, EventArgs e)
        {
            if (this.isRecording && (this.ChangeAction != null))
            {
                HtmlHandler  htmlHandler = (HtmlHandler)sender;
                IHTMLElement htmlSource  = htmlHandler.SourceHTMLWindow.@event.srcElement;
                RecEventArgs recPackage  = RecEventArgs.CreateRecEvent(htmlSource, twebstBrowser);

                // Notify listeners about a recorded value change action.
                this.ChangeAction(this, recPackage);
            }
        }


        // If onmouseup is canceled and onclick is not we'll fall back on OnCanceledClick case.
        private void OnHtmlMouseUp(object sender, EventArgs e)
        {
            // Reset any previous mouse up package.
            this.recLastMouseUpPackage = null;

            // On www.xero.com the source element we get on click event is messed up I don't know why.
            // Just keep the source recorder package we correctly get on mouseup event and use it with onclick event.
            if (this.isRecording && (this.ClickAction != null))
            {
                HtmlHandler  htmlHandler = (HtmlHandler)sender;
                IHTMLElement htmlSource  = htmlHandler.SourceHTMLWindow.@event.srcElement;

                if (!IsValidForClickRec(htmlSource))
                {
                    return;
                }

                // For non <img> elements take a parent anchor. Example: a <b> inside an <a>. Skip anchor with names (page bookmarks).
                IHTMLImgElement imgElem = htmlSource as IHTMLImgElement;
                if (imgElem == null)
                {
                    IElement sourceElement = this.twebstCore.AttachToNativeElement(htmlSource);
                    IElement parentAnchor  = sourceElement.FindParentElement("a", "name="); // A parent anchor without a name.

                    if (parentAnchor != null)
                    {
                        htmlSource = parentAnchor.nativeElement;
                    }
                }

                this.recLastMouseUpPackage = RecEventArgs.CreateRecEvent(htmlSource, twebstBrowser);
            }
        }


        private void OnHtmlClick(object sender, EventArgs e)
        {
            if (this.isRecording && (this.ClickAction != null) && (this.recLastMouseUpPackage != null))
            {
                this.pendingWin32Click = false;

                // On google.com for instance, clicking on a label attached to a radio will trigger another onclick for radio.
                // Duplicate statements must e filtered out.
                if (this.pendingWin32Action &&
                    ((DateTime.Now - this.lastWin32Time).TotalMilliseconds <= MAXIMUM_TIME_FOR_PENDING_CLICKS))
                {
                    this.pendingWin32Action = false;
                    this.lastWin32Time      = DateTime.Now;

                    RecEventArgs recPackage = this.recLastMouseUpPackage;
                    this.recLastMouseUpPackage = null;

                    // Notify listeners about a recorded click action.
                    this.ClickAction(this, recPackage);
                    return;
                }
            }

            this.recLastMouseUpPackage = null;
        }

        
        private bool IsValidForClickRec(IHTMLElement htmlSource)
        {
            String tagName = htmlSource.tagName.ToLower();

            if ((tagName == "textarea") || (tagName == "select") ||
                (tagName == "html")     || (tagName == "body"))
            {
                return false;
            }
            else if (tagName == "input")
            {
                IHTMLInputElement inputElem = (IHTMLInputElement)htmlSource;
                String            type = inputElem.type.ToLower();

                if ((type == "text") || (type == "password") || (type == "file"))
                {
                    return false;
                }
            }

            return true;
        }


        private Recorder()
        {
            this.Visible = false;
        }


        protected override void WndProc (ref Message m)
        {
            if (WM_PENDING_WIN32_CLICK == m.Msg)
            {
                //System.Diagnostics.Trace.WriteLine("WM_PENDING_WIN32_CLICK received");
                if (this.isRecording && this.pendingWin32Click)
                {
                    //System.Diagnostics.Trace.WriteLine("WM_PENDING_WIN32_CLICK to be processed");
                    this.pendingWin32Click = false;

                    Point screenPos = new Point(Win32Api.LoWord(m.WParam.ToInt32()), Win32Api.HiWord(m.WParam.ToInt32()));
                    OnCanceledClick(m.LParam, screenPos);
                }
            }

            base.WndProc(ref m);
        }

        
        private void OnRightClick(IntPtr ieWnd, Point screenPos)
        {
            if (this.isRecording && (this.RightClickAction != null))
            {
                IHTMLElement htmlElem = GetHtmlElementFromScreenPoint(ieWnd, screenPos);
                if (htmlElem == null)
                {
                    return;
                }

                RecEventArgs recPackage = RecEventArgs.CreateRecEvent(htmlElem, twebstBrowser);

                // Notify listeners about a recorded click action.
                this.RightClickAction(this, recPackage);
            }
        }


        private IHTMLElement GetHtmlElementFromScreenPoint(IntPtr ieWnd, Point screenPos)
        {
            if (!Win32Api.IsWindow(ieWnd))
            {
                return null;
            }

            RECT rect = new RECT();

            Win32Api.GetWindowRect(ieWnd, out rect);
            Rectangle screenRectangle = Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);

            if (!screenRectangle.Contains(screenPos))
            {
                return null;
            }

            IAccessible accObj = Win32Api.GetAccessibleObjectFromPoint(screenPos);
            if (accObj == null)
            {
                return null;
            }

            IHTMLElement htmlElem = Win32Api.GetHtmlElementFromAccessible(accObj);
            return htmlElem;
        }


        private void OnCanceledClick(IntPtr ieWnd, Point screenPos)
        {
            // Reset any mouseup package.
            this.recLastMouseUpPackage = null;

            if (!Win32Api.IsWindow(ieWnd))
            {
                return;
            }

            IHTMLElement htmlElem = GetHtmlElementFromScreenPoint(ieWnd, screenPos);
            if (htmlElem == null)
            {
                return;
            }

            if (!IsValidForClickRec(htmlElem))
            {
                return;
            }

            RecEventArgs recPackage = RecEventArgs.CreateRecEvent(htmlElem, twebstBrowser);

            // Notify listeners about a recorded click action.
            this.ClickAction(this, recPackage);
        }


        private static readonly Recorder singleInstance = new Recorder();

        private RecEventArgs    recLastMouseUpPackage           = null;
        private Win32GetMsgHook win32Hook                       = new Win32GetMsgHook();
        private ICore           twebstCore                      = CoreWrapper.Instance;
        private IBrowser        twebstBrowser                   = null;
        private bool            isRecording                     = false;
        private DateTime        lastWin32Time                   = DateTime.Now;
        private bool            pendingWin32Action              = false;
        private bool            pendingWin32Click               = false;
        private const int       MAXIMUM_TIME_FOR_PENDING_CLICKS = 1000;
        private const int       WM_APP                          = 0x8000;
        private const int       WM_PENDING_WIN32_CLICK          = WM_APP + 1;
        
        #endregion
    }
}
