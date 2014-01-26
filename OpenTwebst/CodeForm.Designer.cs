namespace CatStudio
{
    partial class CodeForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeForm));
            this.codeRichTextBox = new System.Windows.Forms.RichTextBox();
            this.codeContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.codeStatusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusProductLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusStateLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLanguageLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLinesLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.saveScriptFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.menuMain = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItemHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.onlineDocumentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reportABugToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.communityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.twebstGoogleGroupToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.twebstLinkedinToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.followUsOnTwitterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.signUpForNewsletterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tutorialsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.samplesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.checkForUpdatesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkForUpdatesBckgWorker = new System.ComponentModel.BackgroundWorker();
            this.codeToolStrip = new CatStudio.ToolStripEx();
            this.codeToolStripBtnNew = new System.Windows.Forms.ToolStripButton();
            this.codeToolStripBtnSave = new System.Windows.Forms.ToolStripButton();
            this.codeToolStripBtnRun = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonRec = new System.Windows.Forms.ToolStripButton();
            this.toolSpyButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonDel = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.codeToolStripLanguageCombo = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.codeContextMenu.SuspendLayout();
            this.codeStatusStrip.SuspendLayout();
            this.menuMain.SuspendLayout();
            this.codeToolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // codeRichTextBox
            // 
            this.codeRichTextBox.AcceptsTab = true;
            this.codeRichTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.codeRichTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.codeRichTextBox.ContextMenuStrip = this.codeContextMenu;
            this.codeRichTextBox.DetectUrls = false;
            this.codeRichTextBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.codeRichTextBox.Location = new System.Drawing.Point(0, 52);
            this.codeRichTextBox.Name = "codeRichTextBox";
            this.codeRichTextBox.ReadOnly = true;
            this.codeRichTextBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth;
            this.codeRichTextBox.ShowSelectionMargin = true;
            this.codeRichTextBox.Size = new System.Drawing.Size(649, 515);
            this.codeRichTextBox.TabIndex = 1;
            this.codeRichTextBox.Text = "";
            this.codeRichTextBox.WordWrap = false;
            // 
            // codeContextMenu
            // 
            this.codeContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.runToolStripMenuItem,
            this.saveToolStripMenuItem});
            this.codeContextMenu.Name = "codeContextMenu";
            this.codeContextMenu.Size = new System.Drawing.Size(165, 70);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.copyToolStripMenuItem.Text = "&Copy         Ctrl+C";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // runToolStripMenuItem
            // 
            this.runToolStripMenuItem.Name = "runToolStripMenuItem";
            this.runToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.runToolStripMenuItem.Text = "&Run Macro";
            this.runToolStripMenuItem.Click += new System.EventHandler(this.runToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.saveToolStripMenuItem.Text = "&Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // codeStatusStrip
            // 
            this.codeStatusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusProductLabel,
            this.toolStripStatusStateLabel,
            this.toolStripStatusLanguageLabel,
            this.toolStripStatusLinesLabel});
            this.codeStatusStrip.Location = new System.Drawing.Point(0, 565);
            this.codeStatusStrip.Name = "codeStatusStrip";
            this.codeStatusStrip.Size = new System.Drawing.Size(649, 25);
            this.codeStatusStrip.TabIndex = 2;
            this.codeStatusStrip.Text = "statusStrip1";
            // 
            // toolStripStatusProductLabel
            // 
            this.toolStripStatusProductLabel.Image = ((System.Drawing.Image)(resources.GetObject("toolStripStatusProductLabel.Image")));
            this.toolStripStatusProductLabel.Name = "toolStripStatusProductLabel";
            this.toolStripStatusProductLabel.Size = new System.Drawing.Size(93, 20);
            this.toolStripStatusProductLabel.Text = "Open Twebst";
            this.toolStripStatusProductLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // toolStripStatusStateLabel
            // 
            this.toolStripStatusStateLabel.AutoSize = false;
            this.toolStripStatusStateLabel.Name = "toolStripStatusStateLabel";
            this.toolStripStatusStateLabel.Size = new System.Drawing.Size(170, 20);
            // 
            // toolStripStatusLanguageLabel
            // 
            this.toolStripStatusLanguageLabel.AutoSize = false;
            this.toolStripStatusLanguageLabel.Name = "toolStripStatusLanguageLabel";
            this.toolStripStatusLanguageLabel.Size = new System.Drawing.Size(60, 20);
            this.toolStripStatusLanguageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // toolStripStatusLinesLabel
            // 
            this.toolStripStatusLinesLabel.Name = "toolStripStatusLinesLabel";
            this.toolStripStatusLinesLabel.Size = new System.Drawing.Size(40, 20);
            this.toolStripStatusLinesLabel.Text = "0 lines";
            // 
            // saveScriptFileDialog
            // 
            this.saveScriptFileDialog.Title = "Save Twebst Macro";
            this.saveScriptFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveScriptFileDialog_FileOk);
            // 
            // menuMain
            // 
            this.menuMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuMain.Location = new System.Drawing.Point(0, 0);
            this.menuMain.Name = "menuMain";
            this.menuMain.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.menuMain.Size = new System.Drawing.Size(649, 24);
            this.menuMain.TabIndex = 4;
            this.menuMain.Text = "Main menu";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.saveToolStripMenuItem1,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(98, 22);
            this.newToolStripMenuItem.Text = "&New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem1
            // 
            this.saveToolStripMenuItem1.Name = "saveToolStripMenuItem1";
            this.saveToolStripMenuItem1.Size = new System.Drawing.Size(98, 22);
            this.saveToolStripMenuItem1.Text = "&Save";
            this.saveToolStripMenuItem1.Click += new System.EventHandler(this.saveToolStripMenuItem1_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(95, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(98, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpToolStripMenuItemHelp,
            this.onlineDocumentationToolStripMenuItem,
            this.samplesToolStripMenuItem,
            this.tutorialsToolStripMenuItem,
            this.toolStripSeparator4,
            this.communityToolStripMenuItem,
            this.reportABugToolStripMenuItem,
            this.toolStripSeparator1,
            this.checkForUpdatesToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // helpToolStripMenuItemHelp
            // 
            this.helpToolStripMenuItemHelp.Name = "helpToolStripMenuItemHelp";
            this.helpToolStripMenuItemHelp.Size = new System.Drawing.Size(196, 22);
            this.helpToolStripMenuItemHelp.Text = "&Help    F1";
            this.helpToolStripMenuItemHelp.Click += new System.EventHandler(this.helpToolStripMenuItemHelp_Click);
            // 
            // onlineDocumentationToolStripMenuItem
            // 
            this.onlineDocumentationToolStripMenuItem.Name = "onlineDocumentationToolStripMenuItem";
            this.onlineDocumentationToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.onlineDocumentationToolStripMenuItem.Text = "Online &Documentation";
            this.onlineDocumentationToolStripMenuItem.Click += new System.EventHandler(this.onlineDocumentationToolStripMenuItem_Click);
            // 
            // reportABugToolStripMenuItem
            // 
            this.reportABugToolStripMenuItem.Name = "reportABugToolStripMenuItem";
            this.reportABugToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.reportABugToolStripMenuItem.Text = "Report a &Bug";
            this.reportABugToolStripMenuItem.Click += new System.EventHandler(this.reportABugToolStripMenuItem_Click);
            // 
            // communityToolStripMenuItem
            // 
            this.communityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.twebstGoogleGroupToolStripMenuItem,
            this.twebstLinkedinToolStripMenuItem,
            this.followUsOnTwitterToolStripMenuItem,
            this.signUpForNewsletterToolStripMenuItem});
            this.communityToolStripMenuItem.Name = "communityToolStripMenuItem";
            this.communityToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.communityToolStripMenuItem.Text = "&Community && Support";
            // 
            // twebstGoogleGroupToolStripMenuItem
            // 
            this.twebstGoogleGroupToolStripMenuItem.Name = "twebstGoogleGroupToolStripMenuItem";
            this.twebstGoogleGroupToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.twebstGoogleGroupToolStripMenuItem.Text = "Twebst Google &Group";
            this.twebstGoogleGroupToolStripMenuItem.Click += new System.EventHandler(this.twebstGoogleGroupToolStripMenuItem_Click);
            // 
            // twebstLinkedinToolStripMenuItem
            // 
            this.twebstLinkedinToolStripMenuItem.Name = "twebstLinkedinToolStripMenuItem";
            this.twebstLinkedinToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.twebstLinkedinToolStripMenuItem.Text = "Twebst &Linkedin Group";
            this.twebstLinkedinToolStripMenuItem.Click += new System.EventHandler(this.twebstLinkedinToolStripMenuItem_Click);
            // 
            // followUsOnTwitterToolStripMenuItem
            // 
            this.followUsOnTwitterToolStripMenuItem.Name = "followUsOnTwitterToolStripMenuItem";
            this.followUsOnTwitterToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.followUsOnTwitterToolStripMenuItem.Text = "Follow us on &Twitter";
            this.followUsOnTwitterToolStripMenuItem.Click += new System.EventHandler(this.followUsOnTwitterToolStripMenuItem_Click);
            // 
            // signUpForNewsletterToolStripMenuItem
            // 
            this.signUpForNewsletterToolStripMenuItem.Name = "signUpForNewsletterToolStripMenuItem";
            this.signUpForNewsletterToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.signUpForNewsletterToolStripMenuItem.Text = "Sign up for &newsletter";
            this.signUpForNewsletterToolStripMenuItem.Click += new System.EventHandler(this.signUpForNewsletterToolStripMenuItem_Click);
            // 
            // tutorialsToolStripMenuItem
            // 
            this.tutorialsToolStripMenuItem.Name = "tutorialsToolStripMenuItem";
            this.tutorialsToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.tutorialsToolStripMenuItem.Text = "&Tutorials";
            this.tutorialsToolStripMenuItem.Click += new System.EventHandler(this.tutorialsToolStripMenuItem_Click);
            // 
            // samplesToolStripMenuItem
            // 
            this.samplesToolStripMenuItem.Name = "samplesToolStripMenuItem";
            this.samplesToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.samplesToolStripMenuItem.Text = "Sam&ples";
            this.samplesToolStripMenuItem.Click += new System.EventHandler(this.samplesToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(193, 6);
            // 
            // checkForUpdatesToolStripMenuItem
            // 
            this.checkForUpdatesToolStripMenuItem.Name = "checkForUpdatesToolStripMenuItem";
            this.checkForUpdatesToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.checkForUpdatesToolStripMenuItem.Text = "Check for &Updates";
            this.checkForUpdatesToolStripMenuItem.Click += new System.EventHandler(this.checkForUpdatesToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // checkForUpdatesBckgWorker
            // 
            this.checkForUpdatesBckgWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.checkForUpdatesBckgWorker_DoWork);
            this.checkForUpdatesBckgWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.checkForUpdatesBckgWorker_RunWorkerCompleted);
            // 
            // codeToolStrip
            // 
            this.codeToolStrip.ClickThrough = false;
            this.codeToolStrip.Dock = System.Windows.Forms.DockStyle.None;
            this.codeToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.codeToolStripBtnNew,
            this.codeToolStripBtnSave,
            this.codeToolStripBtnRun,
            this.toolStripButtonRec,
            this.toolSpyButton,
            this.toolStripButtonDel,
            this.toolStripSeparator3,
            this.codeToolStripLanguageCombo});
            this.codeToolStrip.Location = new System.Drawing.Point(0, 24);
            this.codeToolStrip.Name = "codeToolStrip";
            this.codeToolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.codeToolStrip.Size = new System.Drawing.Size(357, 25);
            this.codeToolStrip.TabIndex = 0;
            this.codeToolStrip.Text = "Code";
            // 
            // codeToolStripBtnNew
            // 
            this.codeToolStripBtnNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.codeToolStripBtnNew.Image = ((System.Drawing.Image)(resources.GetObject("codeToolStripBtnNew.Image")));
            this.codeToolStripBtnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.codeToolStripBtnNew.Name = "codeToolStripBtnNew";
            this.codeToolStripBtnNew.Size = new System.Drawing.Size(23, 22);
            this.codeToolStripBtnNew.Text = "New";
            this.codeToolStripBtnNew.Click += new System.EventHandler(this.codeToolStripBtnNew_Click);
            // 
            // codeToolStripBtnSave
            // 
            this.codeToolStripBtnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.codeToolStripBtnSave.Image = ((System.Drawing.Image)(resources.GetObject("codeToolStripBtnSave.Image")));
            this.codeToolStripBtnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.codeToolStripBtnSave.Name = "codeToolStripBtnSave";
            this.codeToolStripBtnSave.Size = new System.Drawing.Size(23, 22);
            this.codeToolStripBtnSave.Text = "Save";
            this.codeToolStripBtnSave.Click += new System.EventHandler(this.codeToolStripBtnSave_Click);
            // 
            // codeToolStripBtnRun
            // 
            this.codeToolStripBtnRun.Image = ((System.Drawing.Image)(resources.GetObject("codeToolStripBtnRun.Image")));
            this.codeToolStripBtnRun.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.codeToolStripBtnRun.Name = "codeToolStripBtnRun";
            this.codeToolStripBtnRun.Size = new System.Drawing.Size(48, 22);
            this.codeToolStripBtnRun.Text = "Run";
            this.codeToolStripBtnRun.ToolTipText = "Run Macro";
            this.codeToolStripBtnRun.Click += new System.EventHandler(this.codeToolStripBtnRun_Click);
            // 
            // toolStripButtonRec
            // 
            this.toolStripButtonRec.Image = global::CatStudio.Properties.Resources.record_16;
            this.toolStripButtonRec.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonRec.Name = "toolStripButtonRec";
            this.toolStripButtonRec.Size = new System.Drawing.Size(73, 22);
            this.toolStripButtonRec.Text = "Start Rec";
            this.toolStripButtonRec.ToolTipText = "Web Recorder - works only in Twebst browser";
            this.toolStripButtonRec.Click += new System.EventHandler(this.toolStripButtonRec_Click);
            // 
            // toolSpyButton
            // 
            this.toolSpyButton.Image = ((System.Drawing.Image)(resources.GetObject("toolSpyButton.Image")));
            this.toolSpyButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolSpyButton.Name = "toolSpyButton";
            this.toolSpyButton.Size = new System.Drawing.Size(73, 22);
            this.toolSpyButton.Text = "Web Spy";
            this.toolSpyButton.ToolTipText = "Web Spy - works in any IE window";
            this.toolSpyButton.Visible = false;
            this.toolSpyButton.Click += new System.EventHandler(this.toolSpyButton_Click);
            // 
            // toolStripButtonDel
            // 
            this.toolStripButtonDel.Image = global::CatStudio.Properties.Resources.delete_32;
            this.toolStripButtonDel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonDel.Name = "toolStripButtonDel";
            this.toolStripButtonDel.Size = new System.Drawing.Size(60, 22);
            this.toolStripButtonDel.Text = "Delete";
            this.toolStripButtonDel.ToolTipText = "Delete Last Statement";
            this.toolStripButtonDel.Click += new System.EventHandler(this.toolStripButtonDel_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // codeToolStripLanguageCombo
            // 
            this.codeToolStripLanguageCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.codeToolStripLanguageCombo.Name = "codeToolStripLanguageCombo";
            this.codeToolStripLanguageCombo.Size = new System.Drawing.Size(110, 25);
            this.codeToolStripLanguageCombo.SelectedIndexChanged += new System.EventHandler(this.codeToolStripLanguageCombo_SelectedIndexChanged);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(193, 6);
            // 
            // CodeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 590);
            this.Controls.Add(this.menuMain);
            this.Controls.Add(this.codeStatusStrip);
            this.Controls.Add(this.codeRichTextBox);
            this.Controls.Add(this.codeToolStrip);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "CodeForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Open Twebst - Code";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CodeForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CodeForm_FormClosed);
            this.Load += new System.EventHandler(this.CodeForm_Load);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CodeForm_HelpRequested);
            this.codeContextMenu.ResumeLayout(false);
            this.codeStatusStrip.ResumeLayout(false);
            this.codeStatusStrip.PerformLayout();
            this.menuMain.ResumeLayout(false);
            this.menuMain.PerformLayout();
            this.codeToolStrip.ResumeLayout(false);
            this.codeToolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ToolStripEx codeToolStrip;
        private System.Windows.Forms.ToolStripButton codeToolStripBtnNew;
        private System.Windows.Forms.ToolStripButton codeToolStripBtnSave;
        private System.Windows.Forms.ToolStripButton codeToolStripBtnRun;
        private System.Windows.Forms.ToolStripComboBox codeToolStripLanguageCombo;
        private System.Windows.Forms.RichTextBox codeRichTextBox;
        private System.Windows.Forms.StatusStrip codeStatusStrip;
        private System.Windows.Forms.SaveFileDialog saveScriptFileDialog;
        private System.Windows.Forms.ContextMenuStrip codeContextMenu;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuMain;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButtonRec;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItemHelp;
        private System.Windows.Forms.ToolStripMenuItem checkForUpdatesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem onlineDocumentationToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem communityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton toolStripButtonDel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.ComponentModel.BackgroundWorker checkForUpdatesBckgWorker;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusStateLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLanguageLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLinesLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusProductLabel;
        private System.Windows.Forms.ToolStripMenuItem twebstGoogleGroupToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem followUsOnTwitterToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tutorialsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem twebstLinkedinToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem signUpForNewsletterToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolSpyButton;
        private System.Windows.Forms.ToolStripMenuItem samplesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reportABugToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
    }
}