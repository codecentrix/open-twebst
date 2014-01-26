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

using System;
using System.Text;
using System.Windows.Forms;
using mshtml;
using SHDocVw;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;



namespace CatStudio
{
    public partial class BrowserForm : Form
    {
        public BrowserForm()
        {
            InitializeComponent();
            this.Font = System.Drawing.SystemFonts.MessageBoxFont;
            this.codeForm = new CodeForm(this);
        }

        
        public bool IsClosed
        {
            get { return this.closed; }
        }


        private void MainForm_Load(object sender, EventArgs e)
        {
            // Restore position from registry.
            RegistryKey twbstStudioRegKey = Registry.CurrentUser.OpenSubKey("Software\\Codecentrix\\OpenTwebst\\Studio", true);
            bool        showIntroWebPage  = true;

            // If Twebst Studio registry key does not exist then create it.
            if (twbstStudioRegKey == null)
            {
                twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");
            }

            if (twbstStudioRegKey != null)
            {
                // Read position from registry.
                int top    = (int)twbstStudioRegKey.GetValue("mainx", -1);
                int left   = (int)twbstStudioRegKey.GetValue("mainy", -1);
                int height = (int)twbstStudioRegKey.GetValue("mainh", -1);
                int width  = (int)twbstStudioRegKey.GetValue("mainw", -1);

                // Set browser form position and size.
                if ((top != -1) && (left != -1) && (height != -1) && (width != -1) &&
                     IsValidFormPosAndSize(top, left, height, width))
                {
                    this.Top    = top;
                    this.Left   = left;
                    this.Height = height;
                    this.Width  = width;
                }
                else
                {
                    this.GoDefaultPosAndSize();
                }

                int dontShowIntro = (int)twbstStudioRegKey.GetValue("dontshowintro", 0);
                showIntroWebPage  = (dontShowIntro == 0);

                // Close Twebst Studio registry key.
                twbstStudioRegKey.Close();
            }
            else
            {
                this.GoDefaultPosAndSize();
            }

            // Load URL history.
            this.urlHistory.Load();
            foreach (String url in this.urlHistory.Items)
            {
                this.comboBoxAddress.Items.Add(url);
            }

            this.codeForm.Show();

            try
            {
                this.recorder.Init(this.webBrowser);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, CatStudioConstants.TWEBST_PRODUCT_NAME,  MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Environment.Exit(-1);
            }

            // Navigate to intro page if enabled.
            if (showIntroWebPage)
            {
                this.webBrowser.Navigate(CatStudioConstants.INTRO_PAGE_URL);
            }
        }


        private void GoDefaultPosAndSize()
        {
            this.Top    = 0;
            this.Left   = 0;
            this.Height = (2 * Screen.PrimaryScreen.WorkingArea.Height) /3;
            this.Width  = Screen.PrimaryScreen.WorkingArea.Width;
        }


        private bool IsValidFormPosAndSize(int top, int left, int height, int width)
        {
            if ((top < -50) || (left < -50) || (height < 0) || (width < 0))
            {
                return false;
            }

            if (((height - Screen.PrimaryScreen.WorkingArea.Height) > 50)||
                ((width  - Screen.PrimaryScreen.WorkingArea.Width)  > 50))
            {
                return false;
            }

            return true;
        }


        private void GoBtn_Click(object sender, EventArgs e)
        {
            String url = this.comboBoxAddress.Text.Trim();
            if (url.ToLower() == "about:twebst")
            {
                this.webBrowser.Navigate(CatStudioConstants.INTRO_PAGE_URL);
            }
            else
            {
                if (!String.IsNullOrEmpty(url))
                {
                    // Don't add intro page URL to history.
                    if (!IsIntroPage(url))
                    {
                        this.addToHistory = true;
                    }

                    this.webBrowser.Navigate(url);
                }
            }
        }


        private void AddressBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ('\r' == e.KeyChar)
            {
                this.GoBtn_Click(null, null);
            }
        }


        private void buttonFwd_Click(object sender, EventArgs e)
        {
            this.webBrowser.GoForward();

            if (this.recorder.IsRecording)
            {
                this.recorder.RecordForward();
            }
        }


        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.webBrowser.GoBack();

            if (this.recorder.IsRecording)
            {
                this.recorder.RecordBack();
            }
        }


        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.codeForm.CanCloseNow())
            {
                // Save form position in registry.
                RegistryKey twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");

                if (twbstStudioRegKey != null)
                {
                    twbstStudioRegKey.SetValue("mainx", this.Top);
                    twbstStudioRegKey.SetValue("mainy", this.Left);
                    twbstStudioRegKey.SetValue("mainh", this.Height);
                    twbstStudioRegKey.SetValue("mainw", this.Width);

                    twbstStudioRegKey.Close();
                }

                this.urlHistory.Save();
            }
            else
            {
                e.Cancel = true;
            }
        }

        
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.recorder.Uninit();
            this.closed = true;

            if (!this.codeForm.IsClosed)
            {
                this.codeForm.Close();
            }
        }


        private bool IsIntroPage(String url)
        {
            url = url.ToLower();
            return (url == CatStudioConstants.OPEN_INTRO_URL);
        }


        private void webBrowser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            String url = this.webBrowser.Url.ToString();

            if (!IsIntroPage(url))
            {
                this.comboBoxAddress.Text = url;
            }
            else
            {
                this.comboBoxAddress.Text = "about:twebst";
                this.webBrowser.Focus();
                this.HookIntroPage();
            }

            if (this.addToHistory)
            {
                // Don't add about:twebst to url history.
                if (this.comboBoxAddress.Text.ToLower() != "about:twebst")
                {
                    if (this.urlHistory.AddURL(this.comboBoxAddress.Text))
                    {
                        this.comboBoxAddress.Items.Add(this.comboBoxAddress.Text);
                    }
                }

                this.addToHistory = false;
            }
        }


        private void webBrowser_NewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.recorder.IsRecording)
            {
                // Web Recorder does not generate code for multiple browser windows yet, but this scenario IS supported by Twebst Library (see FindBrowser and FindAllBrowsers methods).
                MessageBox.Show(this, "The web site opens a new window!\n\n" +
                                      "Twebst recorder only works in one window.\n"/* + 
                                      "Selection mode works for any web page including IE web control embedded in other apps.\n\n" +
                                      "Switching to selection mode ..."*/,
                                CatStudioConstants.TWEBST_PRODUCT_NAME,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);

                //this.codeForm.StartSelection();
            }
        }


        private void BrowserForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            // Assume that CHM file is in the same dicrectory as the executable.
            String productPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            String chmFilePath = System.IO.Path.Combine(productPath, CatStudioConstants.TWEBST_CHM_FILE_NAME);

            try
            {
                Process process = new Process();
                process.StartInfo.FileName = chmFilePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void comboBoxAddress_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.webBrowser.Focus();
            this.GoBtn_Click(null, null);
        }


        private void HookIntroPage()
        {
            try
            {
                IWebBrowser2   nativeBrowser = (SHDocVw.IWebBrowser2)webBrowser.ActiveXInstance;
                IHTMLDocument3 introDoc      = nativeBrowser.Document as IHTMLDocument3;

                if (introDoc != null)
                {
                    introDoc.attachEvent("onclick", new HtmlHandler(this.OnIntroClick, ((IHTMLDocument2)introDoc).parentWindow));

                    // Initialize checkbox with value in registry.
                    IHTMLInputElement checkBox          = (IHTMLInputElement)introDoc.getElementById(CatStudioConstants.INTRO_DONT_ASK_AGAIN_ATTR);
                    RegistryKey       twbstStudioRegKey = Registry.CurrentUser.OpenSubKey("Software\\Codecentrix\\OpenTwebst\\Studio", false);

                    if (twbstStudioRegKey != null)
                    {
                        int dontShowIntro = (int)twbstStudioRegKey.GetValue("dontshowintro", 0);
                        
                        checkBox.@checked = (dontShowIntro != 0);
                        twbstStudioRegKey.Close();
                    }
                }
            }
            catch
            {
                // Do nothing. Maybe not this critical.
            }
        }


        private void OnIntroClick(object sender, EventArgs e)
        {
            HtmlHandler  htmlHandler = (HtmlHandler)sender;
            IHTMLElement htmlSource  = htmlHandler.SourceHTMLWindow.@event.srcElement;
            String       sourceID    = htmlSource.id;

            if (sourceID == CatStudioConstants.INTRO_START_REC_ATTR)
            {
                this.OnIntroStartRec();
            }
            else if (sourceID == CatStudioConstants.INTRO_RUN_DEMO_ATTR)
            {
                this.OnIntroRunDemo();
            }
            else if (sourceID == CatStudioConstants.INTRO_OPEN_SAMPLES_ATTR)
            {
                this.OnIntroOpenSamples();
            }
            else if (sourceID == CatStudioConstants.INTRO_DONT_ASK_AGAIN_ATTR)
            {
                IHTMLInputElement inputSource = (IHTMLInputElement)htmlSource;
                this.OnIntroDontAskAgain(inputSource.@checked);
            }
        }


        private void OnIntroStartRec()
        {
            this.webBrowser.Navigate(CatStudioConstants.PAGE_STARTED_REC_URL);
            this.codeForm.StartRec();
        }


        private void OnIntroRunDemo()
        {
            try
            {
                String  myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (!String.IsNullOrEmpty(myDocuments))
                {
                    String  sampleDir = Path.Combine(myDocuments, "OpenTwebstSamples\\JScript");
                    if (!Directory.Exists(sampleDir))
                    {
                        throw new Exception();
                    }

                    String  quickTour = Path.Combine(sampleDir, "qt.js");
                    Process process   = new Process();

                    process.StartInfo.FileName        = "wscript.exe";
                    process.StartInfo.Arguments       = "\"" + quickTour + "\"";
                    process.StartInfo.UseShellExecute = true;
                    process.Start();
                }
                else
                {
                    throw new Exception();
                }
            }
            catch
            {
                MessageBox.Show(this, "Can not launch Quick Tour!\nYou may find it in \"My Documents\\OpenTwebst Samples\\qt.js\".",
                                CatStudioConstants.TWEBST_PRODUCT_NAME,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        public void OnIntroOpenSamples()
        {
            try
            {
                String  myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (!String.IsNullOrEmpty(myDocuments))
                {
                    String  sampleDir = Path.Combine(myDocuments, "OpenTwebstSamples");
                    if (!Directory.Exists(sampleDir))
                    {
                        throw new Exception();
                    }

                    Process process     = new Process();
                    process.StartInfo.FileName        = "explorer.exe";
                    process.StartInfo.Arguments       = sampleDir;
                    process.StartInfo.Verb            = "open";
                    process.StartInfo.UseShellExecute = true;
                    process.Start();
                }
                else
                {
                    throw new Exception();
                }
            }
            catch
            {
                MessageBox.Show(this, "Can not open samples!\nYou may find samples in \"My Documents\\OpenTwebstSamples\" folder.",
                                CatStudioConstants.TWEBST_PRODUCT_NAME,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private void OnIntroDontAskAgain(bool isChecked)
        {
            RegistryKey twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");
            if (twbstStudioRegKey != null)
            {
                twbstStudioRegKey.SetValue("dontshowintro", (isChecked ? 1 : 0));
                twbstStudioRegKey.Close();
            }
        }


        // Open help when browser control has focus.
        private void webBrowser_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if ((e.KeyCode == Keys.F1) && this.isKeyDown)
            {
                this.BrowserForm_HelpRequested(null, null);
            }

            this.isKeyDown = !this.isKeyDown;
        }



        // Private data area.
        private Recorder   recorder     = Recorder.Instance;
        private CodeForm   codeForm     = null;
        private bool       closed       = false;
        private URLHistory urlHistory   = new URLHistory();
        private bool       addToHistory = false;
        private bool       isKeyDown    = true;
    }
}
