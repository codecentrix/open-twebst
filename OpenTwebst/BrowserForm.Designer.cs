namespace CatStudio
{
    partial class BrowserForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BrowserForm));
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.comboBoxAddress = new System.Windows.Forms.ComboBox();
            this.toolTipMain = new System.Windows.Forms.ToolTip(this.components);
            this.buttonGo = new System.Windows.Forms.Button();
            this.buttonFwd = new System.Windows.Forms.Button();
            this.buttonBack = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // webBrowser
            // 
            this.webBrowser.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser.IsWebBrowserContextMenuEnabled = false;
            this.webBrowser.Location = new System.Drawing.Point(0, 31);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new System.Drawing.Size(743, 586);
            this.webBrowser.TabIndex = 3;
            this.webBrowser.NewWindow += new System.ComponentModel.CancelEventHandler(this.webBrowser_NewWindow);
            this.webBrowser.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.webBrowser_PreviewKeyDown);
            this.webBrowser.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.webBrowser_Navigated);
            // 
            // comboBoxAddress
            // 
            this.comboBoxAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxAddress.FormattingEnabled = true;
            this.comboBoxAddress.Location = new System.Drawing.Point(54, 3);
            this.comboBoxAddress.Name = "comboBoxAddress";
            this.comboBoxAddress.Size = new System.Drawing.Size(659, 28);
            this.comboBoxAddress.Sorted = true;
            this.comboBoxAddress.TabIndex = 1;
            this.comboBoxAddress.SelectedIndexChanged += new System.EventHandler(this.comboBoxAddress_SelectedIndexChanged);
            this.comboBoxAddress.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.AddressBox_KeyPress);
            // 
            // buttonGo
            // 
            this.buttonGo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonGo.AutoSize = true;
            this.buttonGo.Image = global::CatStudio.Properties.Resources.go_16;
            this.buttonGo.Location = new System.Drawing.Point(717, 2);
            this.buttonGo.Name = "buttonGo";
            this.buttonGo.Size = new System.Drawing.Size(25, 23);
            this.buttonGo.TabIndex = 2;
            this.toolTipMain.SetToolTip(this.buttonGo, "Go");
            this.buttonGo.UseVisualStyleBackColor = true;
            this.buttonGo.Click += new System.EventHandler(this.GoBtn_Click);
            // 
            // buttonFwd
            // 
            this.buttonFwd.AutoSize = true;
            this.buttonFwd.Image = global::CatStudio.Properties.Resources.arrow_forward_16;
            this.buttonFwd.Location = new System.Drawing.Point(26, 3);
            this.buttonFwd.Name = "buttonFwd";
            this.buttonFwd.Size = new System.Drawing.Size(25, 24);
            this.buttonFwd.TabIndex = 5;
            this.toolTipMain.SetToolTip(this.buttonFwd, "Navigate Forward");
            this.buttonFwd.UseVisualStyleBackColor = true;
            this.buttonFwd.Click += new System.EventHandler(this.buttonFwd_Click);
            // 
            // buttonBack
            // 
            this.buttonBack.AutoSize = true;
            this.buttonBack.Image = global::CatStudio.Properties.Resources.arrow_back_16;
            this.buttonBack.Location = new System.Drawing.Point(1, 3);
            this.buttonBack.Name = "buttonBack";
            this.buttonBack.Size = new System.Drawing.Size(25, 24);
            this.buttonBack.TabIndex = 4;
            this.toolTipMain.SetToolTip(this.buttonBack, "Navigate Back");
            this.buttonBack.UseVisualStyleBackColor = true;
            this.buttonBack.Click += new System.EventHandler(this.buttonBack_Click);
            // 
            // BrowserForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(743, 618);
            this.Controls.Add(this.buttonGo);
            this.Controls.Add(this.comboBoxAddress);
            this.Controls.Add(this.buttonFwd);
            this.Controls.Add(this.buttonBack);
            this.Controls.Add(this.webBrowser);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "BrowserForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Open Twebst - Browser";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.BrowserForm_HelpRequested);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Button buttonBack;
        private System.Windows.Forms.Button buttonFwd;
        private System.Windows.Forms.ComboBox comboBoxAddress;
        private System.Windows.Forms.ToolTip toolTipMain;
        private System.Windows.Forms.Button buttonGo;
    }
}

