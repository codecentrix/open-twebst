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
using System.Diagnostics;


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
        public event EventHandler<RecEventArgs> ElementSelected;

        public static Recorder Instance
        {
            get
            {
                return singleInstance;
            }
        }


        public void StartRec()
        {
            if (this.recorderMode == RecorderMode.RECORDING_MODE)
            {
                // Already recording.
                return;
            }
            else if (this.recorderMode == RecorderMode.SELECTION_MODE)
            {
                StopSelection();
            }

            Debug.Assert(this.recorderMode == RecorderMode.STOP_MODE);
            this.recorderMode = RecorderMode.RECORDING_MODE;

            if (RecStarted != null)
            {
                RecStarted(this, EventArgs.Empty);
            }
        }


        public void StopRec()
        {
            if (this.recorderMode != RecorderMode.RECORDING_MODE)
            {
                // Not recording.
                return;
            }

            this.recorderMode = RecorderMode.STOP_MODE;
            if (RecStopped != null)
            {
                RecStopped(this, EventArgs.Empty);
            }
        }


        public bool IsRecording
        {
            get { return this.recorderMode == RecorderMode.RECORDING_MODE; }
        }


        public bool IsSelecting
        {
            get { return this.recorderMode == RecorderMode.SELECTION_MODE; }
        }

        public void StartSelection()
        {
            if (this.recorderMode == RecorderMode.RECORDING_MODE)
            {
                StopRec();
            }

            Debug.Assert(this.recorderMode == RecorderMode.STOP_MODE);
            this.recorderMode = RecorderMode.SELECTION_MODE;
            this.win32MouseHook.Install();
        }


        public void StopSelection()
        {
            if (this.recorderMode != RecorderMode.SELECTION_MODE)
            {
                // Not in selection mode.
                return;
            }

            this.recorderMode = RecorderMode.STOP_MODE;
            this.win32MouseHook.UnInstall();
            RemoveHighlight();
        }


        public void RecordForward()
        {
            if (this.IsRecording && (this.ForwardAction != null))
            {
                this.ForwardAction(this, EventArgs.Empty);
            }
        }


        public void RecordBack()
        {
            if (this.IsRecording && (this.BackAction != null))
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

            this.win32MouseHook.Win32HookMouseMsg += OnMouseMsg;
        }


        public void Uninit()
        {
            this.win32Hook.Win32HookMsg -= OnWin32HookMsg;
            this.win32Hook.UnInstall();

            this.win32MouseHook.Win32HookMouseMsg -= OnMouseMsg;
            this.win32MouseHook.UnInstall();

            this.StopSelection();
            this.twebstBrowser = null;
            this.twebstCore.Reset();
        }
        #endregion

        #region Private section
        private void OnMouseMsg(Object sender, Win32MouseLLEventArgs e)
        {
            // The hook procedure should process a message in less than a registry specified time
            // otherwise the OS will silently removed the hook. The registry value is 5 seconds.
            if (e.Message == Win32Api.WM_LBUTTONUP)
            {

                Win32Api.PostMessage(this.Handle, WM_DELAYED_SELECTED, e.ScreenPoint.X, e.ScreenPoint.Y);
            }
            else if (e.Message == Win32Api.WM_MOUSEMOVE)
            {
                Win32Api.PostMessage(this.Handle, WM_DELAYED_SELECTING, e.ScreenPoint.X, e.ScreenPoint.Y);
            }
        }

        private void OnWin32HookMsg(Object sender, Win32HookMsgEventArgs e)
        {
            if (this.IsRecording)
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
            if (this.IsRecording && (this.ChangeAction != null))
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
            if (this.IsRecording && (this.ClickAction != null))
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
            if (this.IsRecording && (this.ClickAction != null) && (this.recLastMouseUpPackage != null))
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
                if (this.IsRecording && this.pendingWin32Click)
                {
                    //System.Diagnostics.Trace.WriteLine("WM_PENDING_WIN32_CLICK to be processed");
                    this.pendingWin32Click = false;

                    Point screenPos = new Point(Win32Api.LoWord(m.WParam.ToInt32()), Win32Api.HiWord(m.WParam.ToInt32()));
                    OnCanceledClick(m.LParam, screenPos);
                }
            }
            else if (WM_DELAYED_SELECTED == m.Msg)
            {
                if (this.IsSelecting)
                {
                    int x = m.WParam.ToInt32();
                    int y = m.LParam.ToInt32();
                    OnScreenSelected(x, y);
                }
            }
            else if (WM_DELAYED_SELECTING == m.Msg)
            {
                if (this.IsSelecting)
                {
                    int x = m.WParam.ToInt32();
                    int y = m.LParam.ToInt32();
                    OnScreenSelecting(x, y);
                }
            }

            base.WndProc(ref m);
        }


        private void OnScreenSelected(int x, int y)
        {
            Debug.Assert(this.IsSelecting);

            if (this.ElementSelected != null)
            {
                try
                {
                    IElement elem = this.lastSelectingElem; //this.twebstCore.FindElementFromPoint(x, y);
                    if (elem != null)
                    {
                        IHTMLElement htmlElem = elem.nativeElement;
                        RecEventArgs recPackage = RecEventArgs.CreateRecEvent(htmlElem, elem.parentBrowser);
                        this.ElementSelected(this, recPackage);
                    }
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            }
        }


        private void RemoveHighlight()
        {
            try
            {
                if (this.lastSelectingElem != null)
                {
                    // Restore background.
                    this.lastSelectingElem.nativeElement.style.backgroundColor = this.savedSelectBackground;

                    if (ie8orLater)
                    {
                        ((IHTMLStyle6)(this.lastSelectingElem.nativeElement.style)).outline = this.savedSelectOutline;
                    }

                    this.lastSelectingElem.RemoveAttribute(CatStudioConstants.CRNT_SELECTION_ATTR);

                    // Cleanup.
                    Marshal.ReleaseComObject(this.lastSelectingElem);
                    this.lastSelectingElem     = null;
                    this.savedSelectBackground = null;
                    this.savedSelectOutline    = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                this.lastSelectingElem = null;
            }
        }


        private void PutHighlight(IElement elem)
        {
            Debug.Assert(this.lastSelectingElem == null);

            // Set the background.
            IHTMLElement2     elem2 = (IHTMLElement2)(elem.nativeElement);
            IHTMLCurrentStyle crntStyle = elem2.currentStyle;

            this.savedSelectBackground = crntStyle.backgroundColor;
            elem.nativeElement.style.backgroundColor = CatStudioConstants.TWEBST_SELECT_BCKG;

            // outline style is available starting with IE8.
            if (ie8orLater)
            {
                IHTMLStyle6 style = (IHTMLStyle6)(elem.nativeElement.style);
                IHTMLCurrentStyle5 crntStyle5 = (IHTMLCurrentStyle5)crntStyle;

                this.savedSelectOutline = crntStyle5.outline;
                style.outline = CatStudioConstants.TWEBST_SELECT_OUTLINE;
            }

            elem.SetAttribute(CatStudioConstants.CRNT_SELECTION_ATTR, "1");
            this.lastSelectingElem  = elem;
        }


        private void OnScreenSelecting(int x, int y)
        {
            Debug.Assert(this.IsSelecting);

            if (this.ElementSelected != null)
            {
                try
                {
                    IElement elem = this.twebstCore.FindElementFromPoint(x, y);
                    if (elem != null)
                    {
                        // Check if new element.
                        if (String.IsNullOrEmpty(((String)elem.GetAttribute(CatStudioConstants.CRNT_SELECTION_ATTR))))
                        {
                            // Remove selection highlight on old elem.
                            RemoveHighlight();
                            PutHighlight(elem);
                        }
                    }
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            }
        }

        
        private void OnRightClick(IntPtr ieWnd, Point screenPos)
        {
            if (this.IsRecording && (this.RightClickAction != null))
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

        private enum RecorderMode
        {
            STOP_MODE,
            RECORDING_MODE,
            SELECTION_MODE,
        }

        private static readonly Recorder singleInstance = new Recorder();

        private RecEventArgs     recLastMouseUpPackage           = null;
        private Win32GetMsgHook  win32Hook                       = new Win32GetMsgHook();
        private Win32LLMouseHook win32MouseHook                  = new Win32LLMouseHook();
        private ICore            twebstCore                      = CoreWrapper.Instance;
        private bool             ie8orLater                      = (CoreWrapper.Instance.IEVersion.CompareTo("8") >= 0);
        private IBrowser         twebstBrowser                   = null;
        private RecorderMode     recorderMode                    = RecorderMode.STOP_MODE;
        private DateTime         lastWin32Time                   = DateTime.Now;
        private IElement         lastSelectingElem               = null;
        private Object           savedSelectBackground           = null;
        private String           savedSelectOutline              = null;
        private bool             pendingWin32Action              = false;
        private bool             pendingWin32Click               = false;
        private const int        MAXIMUM_TIME_FOR_PENDING_CLICKS = 1000;
        private const int        WM_APP                          = 0x8000;
        private const int        WM_PENDING_WIN32_CLICK          = WM_APP + 1;
        private const int        WM_DELAYED_SELECTED             = WM_APP + 2;
        private const int        WM_DELAYED_SELECTING            = WM_APP + 3;

        #endregion
    }
}
