namespace CatStudio
{
    partial class AboutBox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutBox));
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            this.okButton = new System.Windows.Forms.Button();
            this.labelProduct = new System.Windows.Forms.Label();
            this.labelVersion = new System.Windows.Forms.Label();
            this.linkLabelCodecentrix = new System.Windows.Forms.LinkLabel();
            this.labelGpl = new System.Windows.Forms.Label();
            this.linkLabelMail = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // logoPictureBox
            // 
            this.logoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("logoPictureBox.Image")));
            this.logoPictureBox.Location = new System.Drawing.Point(14, 14);
            this.logoPictureBox.Name = "logoPictureBox";
            this.logoPictureBox.Size = new System.Drawing.Size(76, 81);
            this.logoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.logoPictureBox.TabIndex = 25;
            this.logoPictureBox.TabStop = false;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Location = new System.Drawing.Point(324, 198);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(87, 30);
            this.okButton.TabIndex = 1;
            this.okButton.Text = "&OK";
            // 
            // labelProduct
            // 
            this.labelProduct.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProduct.Location = new System.Drawing.Point(101, 14);
            this.labelProduct.Name = "labelProduct";
            this.labelProduct.Size = new System.Drawing.Size(304, 25);
            this.labelProduct.TabIndex = 27;
            this.labelProduct.Text = "Product name";
            // 
            // labelVersion
            // 
            this.labelVersion.AutoSize = true;
            this.labelVersion.Location = new System.Drawing.Point(102, 48);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(46, 15);
            this.labelVersion.TabIndex = 28;
            this.labelVersion.Text = "Version";
            // 
            // linkLabelCodecentrix
            // 
            this.linkLabelCodecentrix.AutoSize = true;
            this.linkLabelCodecentrix.Location = new System.Drawing.Point(13, 206);
            this.linkLabelCodecentrix.Name = "linkLabelCodecentrix";
            this.linkLabelCodecentrix.Size = new System.Drawing.Size(247, 15);
            this.linkLabelCodecentrix.TabIndex = 32;
            this.linkLabelCodecentrix.TabStop = true;
            this.linkLabelCodecentrix.Text = "https://github.com/codecentrix/open-twebst";
            this.linkLabelCodecentrix.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelCodecentrix_LinkClicked);
            // 
            // labelGpl
            // 
            this.labelGpl.AutoSize = true;
            this.labelGpl.Location = new System.Drawing.Point(13, 115);
            this.labelGpl.Name = "labelGpl";
            this.labelGpl.Size = new System.Drawing.Size(298, 15);
            this.labelGpl.TabIndex = 28;
            this.labelGpl.Text = "This program is released under the terms of the GPL v3.";
            // 
            // linkLabelMail
            // 
            this.linkLabelMail.AutoSize = true;
            this.linkLabelMail.Location = new System.Drawing.Point(13, 182);
            this.linkLabelMail.Name = "linkLabelMail";
            this.linkLabelMail.Size = new System.Drawing.Size(147, 15);
            this.linkLabelMail.TabIndex = 33;
            this.linkLabelMail.TabStop = true;
            this.linkLabelMail.Text = "support@codecentrix.com";
            this.linkLabelMail.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMail_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(102, 80);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 15);
            this.label1.TabIndex = 34;
            this.label1.Text = "Codecentrix 2014";
            // 
            // AboutBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.okButton;
            this.ClientSize = new System.Drawing.Size(426, 237);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.linkLabelMail);
            this.Controls.Add(this.linkLabelCodecentrix);
            this.Controls.Add(this.labelGpl);
            this.Controls.Add(this.labelVersion);
            this.Controls.Add(this.labelProduct);
            this.Controls.Add(this.logoPictureBox);
            this.Controls.Add(this.okButton);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutBox";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Open Twebst - About";
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox logoPictureBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label labelProduct;
        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.LinkLabel linkLabelCodecentrix;
        private System.Windows.Forms.Label labelGpl;
        private System.Windows.Forms.LinkLabel linkLabelMail;
        private System.Windows.Forms.Label label1;



    }
}
