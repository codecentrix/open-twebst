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
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
using System.Xml;



namespace CatStudio
{
    public partial class CodeForm : Form
    {
        public CodeForm(BrowserForm bf)
        {
            InitializeComponent();
            this.Font        = System.Drawing.SystemFonts.MessageBoxFont;
            this.browserForm = bf;
        }

        
        public bool IsClosed
        {
            get { return this.closed; }
        }


        public void StartRec()
        {
            if (!Recorder.Instance.IsRecording)
            {
                this.toolStripButtonRec_Click(null, null);
            }
        }


        public void StartSelection()
        {
            if (!Recorder.Instance.IsSelecting)
            {
                this.toolSpyButton_Click(null, null);
            }
        }

        private void CodeForm_Load(object sender, EventArgs e)
        {
            // Restore position from registry.
            RegistryKey twbstStudioRegKey = Registry.CurrentUser.OpenSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");

            // If Twebst Studio registry key does not exist then create it.
            if (twbstStudioRegKey == null)
            {
                twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");
            }

            if (twbstStudioRegKey != null)
            {
                int top    = (int)twbstStudioRegKey.GetValue("codex", -1);
                int left   = (int)twbstStudioRegKey.GetValue("codey", -1);
                int height = (int)twbstStudioRegKey.GetValue("codeh", -1);
                int width  = (int)twbstStudioRegKey.GetValue("codew", -1);

                if ((top != -1) && (left != -1) && (height != -1) && (width != -1) &&
                     IsValidFormPosAndSize(top, left, height, width))
                {
                    this.Top = top;
                    this.Left = left;
                    this.Height = height;
                    this.Width = width;
                }
                else
                {
                    this.GoDefaultPosAndSize();
                }

                twbstStudioRegKey.Close();
            }
            else
            {
                this.GoDefaultPosAndSize();
            }

            // Create target language objects.
            BaseLanguageGenerator vbsLang    = new VbsGenerator();
            BaseLanguageGenerator jsLang     = new JsGenerator();
            BaseLanguageGenerator pythonLang = new PyGenerator();
            BaseLanguageGenerator watirEnv   = new WatirGenerator();
            BaseLanguageGenerator csLang     = new CSharpGenerator();
            BaseLanguageGenerator vbNetLang  = new VbNetGenerator();
            BaseLanguageGenerator vbaLang    = new VbaGenerator();

            // Populate language combo-box.
            this.codeToolStripLanguageCombo.Items.Add(vbsLang);
            this.codeToolStripLanguageCombo.Items.Add(jsLang);
            this.codeToolStripLanguageCombo.Items.Add(vbaLang);
            this.codeToolStripLanguageCombo.Items.Add(csLang);
            this.codeToolStripLanguageCombo.Items.Add(vbNetLang);
            this.codeToolStripLanguageCombo.Items.Add(pythonLang);
            this.codeToolStripLanguageCombo.Items.Add(watirEnv);

            // Initialize code generator.
            this.codeGen = new CodeGenerator(vbsLang);
            this.codeGen.NewStatement += OnNewStatement;
            this.codeGen.CodeChanged  += OnLanguageChanged;

            // Select a language in the combo-box.
            int defaultLanguage = GetRecorderSavedLanguage();
            if ((defaultLanguage < 0) || (defaultLanguage >= this.codeToolStripLanguageCombo.Items.Count))
            {
                defaultLanguage = 0;
            }

            this.codeToolStripLanguageCombo.SelectedIndex = defaultLanguage;

            this.codeToolStrip.ClickThrough = true;
            this.toolStripStatusProductLabel.Text = CoreWrapper.Instance.productName + " " + CoreWrapper.Instance.productVersion;
        }


        private int GetRecorderSavedLanguage()
        {
            RegistryKey twbstStudioRegKey = null;
            int language = 0;

            try
            {
                twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");
                language = (int)twbstStudioRegKey.GetValue("recorderLanguage", 0);
            }
            finally
            {
                if (twbstStudioRegKey != null)
                {
                    twbstStudioRegKey.Close();
                }
            }

            return language;
        }


        private void SaveRecorderLanguage(int language)
        {
            RegistryKey twbstStudioRegKey = null;

            try
            {
                twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");
                twbstStudioRegKey.SetValue("recorderLanguage", language);
            }
            finally
            {
                if (twbstStudioRegKey != null)
                {
                    twbstStudioRegKey.Close();
                }
            }
        }


        private void GoDefaultPosAndSize()
        {
            this.Top    = (2 * Screen.PrimaryScreen.WorkingArea.Height) /3;
            this.Left   = 0;
            this.Height = Screen.PrimaryScreen.WorkingArea.Height /3;
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

        private void CodeForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.closed = true;

            if (!this.browserForm.IsClosed)
            {
                this.browserForm.Close();
            }
        }


        private void CodeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!this.CanCloseNow())
            {
                // Don't exit
                e.Cancel = true;
            }
            else
            {
                // Save form position in registry.
                RegistryKey twbstStudioRegKey = Registry.CurrentUser.CreateSubKey("Software\\Codecentrix\\OpenTwebst\\Studio");

                if (twbstStudioRegKey != null)
                {
                    twbstStudioRegKey.SetValue("codex", this.Top);
                    twbstStudioRegKey.SetValue("codey", this.Left);
                    twbstStudioRegKey.SetValue("codeh", this.Height);
                    twbstStudioRegKey.SetValue("codew", this.Width);

                    twbstStudioRegKey.Close();
                }
            }
        }


        public bool CanCloseNow()
        {
            if (this.canCloseNow || !this.isDirty)
            {
                // Already asked.
                return true;
            }

            if (this.codeGen.GetAllCode().Length > 0)
            {
                DialogResult answer = MessageBox.Show(this, "Do you want to save changes?",
                                                      CatStudioConstants.TWEBST_PRODUCT_NAME,
                                                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (answer == DialogResult.No)
                {
                    canCloseNow = true;
                    return true;
                }
                else if (answer == DialogResult.Yes)
                {
                    // Save the macro.
                    return (canCloseNow = SaveScript());
                }
                else
                {
                    // Cancel button. Don't close.
                    canCloseNow = false;
                    return false;
                }
            }
            else
            {
                canCloseNow = true;
                return true;
            }
        }


        private void OnLanguageChanged(Object sender, CodeGenEventArgs args)
        {
            if (args.Code != "")
            {
                this.SetRichEditBoxText(this.codeGen.GetAllDecoratedCode());
                this.isDirty = true;
            }
            else
            {
                this.SetRichEditBoxText("");
            }
        }


        private void OnNewStatement(Object sender, CodeGenEventArgs args)
        {
            String code = this.codeGen.GetAllDecoratedCode();

            this.SetRichEditBoxText(code);
            this.isDirty = true;
        }


        private void codeToolStripLanguageCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            BaseLanguageGenerator targetLang    = (BaseLanguageGenerator)this.codeToolStripLanguageCombo.SelectedItem;
            bool                  resetRecorder = ((this.crntLang is WatirGenerator) || (targetLang is WatirGenerator));

            this.crntLang = targetLang;

            if (resetRecorder)
            {
                // Reset the recorder.
                this.codeGen.Reset();
                this.isDirty = false;

                if (targetLang != null)
                {
                    this.codeGen.Language = targetLang;
                }

                this.SetRichEditBoxText("Switching between Twebst <-> Watir is not supported!\nRecording session restarted ...");
            }
            else
            {
                if (targetLang != null)
                {
                    this.codeGen.Language = targetLang;
                }

                this.toolStripStatusLanguageLabel.Text = targetLang.ToString();
                this.codeRichTextBox.Focus();
                this.codeRichTextBox.Select(0, 0);
                this.codeRichTextBox.ScrollToCaret();
            }

            SaveRecorderLanguage(this.codeToolStripLanguageCombo.SelectedIndex);
        }


        private void codeToolStripBtnRun_Click(object sender, EventArgs e)
        {
            String code = this.codeGen.GetAllDecoratedCode();

            if (code != "")
            {
                try
                {
                    this.codeGen.Language.Play(code);
                }
                catch (System.Reflection.TargetInvocationException ex)
                {
                    MessageBox.Show(this, "Twebst Library error: '" + ex.InnerException.Message + "'", CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(this, "Nothing to play! Record some actions first.", CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void codeToolStripBtnNew_Click(object sender, EventArgs e)
        {
            if (this.codeGen.GetAllCode().Length > 0)
            {
                DialogResult answer = MessageBox.Show(this, "Do you want to save changes?", CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (answer == DialogResult.Cancel)
                {
                    return;
                }
                else if (answer == DialogResult.Yes)
                {
                    // Save the macro.
                    bool saved = SaveScript();

                    if (!saved)
                    {
                        // The user canceled the saving.
                        return;
                    }
                }

                // Reset the recorder.
                this.codeGen.Reset();
                this.SetRichEditBoxText("");
                this.isDirty = false;
            }
        }


        private void codeToolStripBtnSave_Click(object sender, EventArgs e)
        {
            if (this.codeGen.GetAllCode().Length > 0)
            {
                SaveScript();
            }
            else
            {
                MessageBox.Show(this, "Nothing to save! Record some actions first.", CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private bool SaveScript()
        {
            BaseLanguageGenerator targetLang = this.codeGen.Language;

            this.saveScriptFileDialog.Filter     = targetLang.ToString() + " files (*" + targetLang.FileExt + ")|*" + targetLang.FileExt;
            this.saveScriptFileDialog.DefaultExt = targetLang.FileExt.Remove(0, 1);

            DialogResult answer = this.saveScriptFileDialog.ShowDialog(this);
            if (answer == DialogResult.OK)
            {
                // Save the file.
                String     fileName   = this.saveScriptFileDialog.FileName;
                TextWriter textWriter = new StreamWriter(fileName, false, targetLang.GeneratorEncoding);

                textWriter.Write(this.codeGen.GetAllDecoratedCode());
                textWriter.Close();

                this.saveScriptFileDialog.FileName = "";
                this.isDirty = false;

                return true;
            }
            else
            {
                return false;
            }
        }


        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox     about = new AboutBox();
            DialogResult res   = about.ShowDialog(this);
        }


        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            codeToolStripBtnRun_Click(sender, e);
        }


        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            codeToolStripBtnSave_Click(sender, e);
        }


        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.codeRichTextBox.SelectionLength > 0)
            {
                this.codeRichTextBox.Copy();
            }
            else
            {
                this.codeRichTextBox.SelectAll();
                this.codeRichTextBox.Copy();
                this.codeRichTextBox.DeselectAll();
            }
        }


        private void RecordingUI(bool recorderOn)
        {
            if (recorderOn)
            {
                SelectionUI(false);

                this.toolStripButtonRec.Checked     = true;
                this.toolStripButtonRec.Text        = "Stop Rec";
                this.toolStripStatusStateLabel.Text = "Recording...";
            }
            else
            {
                this.toolStripButtonRec.Checked     = false;
                this.toolStripButtonRec.Text        = "Start Rec";
                this.toolStripStatusStateLabel.Text = "";
            }
        }

        private void toolStripButtonRec_Click(object sender, EventArgs e)
        {
            if (!Recorder.Instance.IsRecording)
            {
                RecordingUI(true);
                Recorder.Instance.StartRec();
            }
            else
            {
                RecordingUI(false);
                Recorder.Instance.StopRec();
            }
        }


        private void saveScriptFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            String selectedFileExt = Path.GetExtension(this.saveScriptFileDialog.FileName).ToUpper();
            String expectedFileExt = this.codeGen.Language.FileExt.ToUpper();

            if (selectedFileExt != expectedFileExt)
            {
                e.Cancel = true;
                MessageBox.Show(this, "Invalid file extension!\n" + expectedFileExt + " is expected.", CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void helpToolStripMenuItemHelp_Click(object sender, EventArgs e)
        {
            OpenURL(GetHelpFilePath());
        }


        private void onlineDocumentationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_ONLINE_HELP_URL);
        }


        private void OpenURL(String url)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = url;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private String GetHelpFilePath()
        {
            // Assume that CHM file is in the same dicrectory as the executable.
            String productPath = Path.GetDirectoryName(Application.ExecutablePath);
            String chmFilePath = Path.Combine(productPath, CatStudioConstants.TWEBST_CHM_FILE_NAME);

            return chmFilePath;
        }


        private void CodeForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            OpenURL(GetHelpFilePath());
        }


        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.codeToolStripBtnNew_Click(sender, e);
        }


        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.saveToolStripMenuItem_Click(sender, e);
        }


        private void toolStripButtonDel_Click(object sender, EventArgs e)
        {
            this.codeGen.DeleteLastStatement();
        }

        
        private void checkForUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.toolStripStatusStateLabel.Text = "Checking for Updates...";
            this.checkForUpdatesBckgWorker.RunWorkerAsync();
        }


        private void checkForUpdatesBckgWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.Load(CatStudioConstants.TWEBST_CHECK_UPDATES_URL);
                
                String lastSiteVer = xmlDoc.DocumentElement.GetAttribute(CatStudioConstants.CHECK_UPDATES_VERSION_ATTR);
                String downloadURL = xmlDoc.DocumentElement.GetAttribute(CatStudioConstants.CHECK_UPDATES_DOWNLOAD_URL_ATTR);

                if (lastSiteVer.CompareTo(CoreWrapper.Instance.productVersion) > 0)
                {
                    String                msg = String.Format("A new version ({0}) of {1} is available at {2}\n\nDo you want to download it now?", lastSiteVer, CatStudioConstants.TWEBST_PRODUCT_NAME, downloadURL);
                    CheckForUpdatesResult res = new CheckForUpdatesResult(true, false, msg);

                    res.downloadUrl = downloadURL;
                    e.Result        = res;
                }
                else
                {
                    String msg = String.Format("{0} is up to date ({1})", CatStudioConstants.TWEBST_PRODUCT_NAME, CoreWrapper.Instance.productVersion);
                    e.Result = new CheckForUpdatesResult(true, true, msg);
                }
            }
            catch (Exception ex)
            {
                e.Result = new CheckForUpdatesResult(false, false, "Check for updates failed: " + ex.Message);
            }
        }


        private void checkForUpdatesBckgWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (Recorder.Instance.IsRecording)
            {
                this.toolStripStatusStateLabel.Text = "Recording...";
            }
            else
            {
                this.toolStripStatusStateLabel.Text = "";
            }

            CheckForUpdatesResult res = (CheckForUpdatesResult)e.Result;

            if (!res.succeeded)
            {
                // An error occured or product is up to date.
                MessageBox.Show(this, res.messsage, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (res.isUpToDate)
                {
                    MessageBox.Show(this, res.messsage, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    DialogResult dlgRes = MessageBox.Show(this, res.messsage, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (DialogResult.Yes == dlgRes)
                    {
                        this.OpenURL(res.downloadUrl);
                    }
                }
            }
        }


        private void SetRichEditBoxText(String text)
        {
            this.codeRichTextBox.Text = text;
            this.codeRichTextBox.Select(text.Length, 0);
            this.codeRichTextBox.ScrollToCaret();
            this.toolStripStatusLinesLabel.Text = String.Format("{0} lines", this.codeRichTextBox.Lines.Length);
        }


        private void twebstGoogleGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_COMMUNITY_URL);
        }


        private void followUsOnTwitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_TWITTER_URL);
        }


        private void tutorialsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_TUTORIALS_URL);
        }


        private void twebstLinkedinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_LINKEDIN_URL);
        }


        private void signUpForNewsletterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenURL(CatStudioConstants.TWEBST_NEWSLETTER_URL);
        }


        private void SelectionUI(bool selectionOn)
        {
            if (selectionOn)
            {
                RecordingUI(false);
                this.toolSpyButton.Checked = true;
                this.toolStripStatusStateLabel.Text = "Selection...";
            }
            else
            {
                this.toolSpyButton.Checked = false;
                this.toolStripStatusStateLabel.Text = "";
            }
        }


        private void toolSpyButton_Click(object sender, EventArgs e)
        {
            if (!Recorder.Instance.IsSelecting)
            {
                SelectionUI(true);
                Recorder.Instance.StartSelection();
            }
            else
            {
                SelectionUI(false);
                Recorder.Instance.StopSelection();
            }
        }


        private void samplesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.browserForm.OnIntroOpenSamples();
        }


        private void reportABugToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SupportSystemInfo si = new SupportSystemInfo();
            String supportBody = CatStudioConstants.SUPPORT_EMAIL_MESSAGE + "%0a%0a%0a" + si.Info;
            String supportSubj = CatStudioConstants.SUPPORT_EMAIL_SUBJECT;
            String supportCmd  = String.Format("mailto:support@codecentrix.com?subject={0}&body={1}", supportSubj, supportBody);

            OpenURL(supportCmd);
        }

        private BaseLanguageGenerator crntLang    = null;
        private CodeGenerator          codeGen     = null;
        private BrowserForm            browserForm = null;
        private bool                   canCloseNow = false;
        private bool                   closed      = false;
        private bool                   isDirty     = false;
    }



    internal class CheckForUpdatesResult
    {
        public CheckForUpdatesResult(bool succeeded, bool isUpToDate, String messsage)
        {
            this.succeeded  = succeeded;
            this.isUpToDate = isUpToDate;
            this.messsage   = messsage;
        }

        internal bool   isUpToDate;
        internal bool   succeeded;
        internal String messsage;
        internal String downloadUrl = null;
    }
}
