using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;



namespace CatStudio
{
    public partial class DemoWarning : Form
    {
        public DemoWarning()
        {
            InitializeComponent();
            this.Font = System.Drawing.SystemFonts.MessageBoxFont;
        }


        private void buttonBuy_Click(object sender, EventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = CatStudioConstants.BUY_NOW_URL;
            process.StartInfo.UseShellExecute = true;
            process.Start();

            this.Close();
        }


        private void buttonRegister_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
