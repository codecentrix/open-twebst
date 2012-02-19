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
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;



namespace CatStudio
{
    partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();

            this.Font              = System.Drawing.SystemFonts.MessageBoxFont;
            this.labelProduct.Text = CoreWrapper.Instance.productName.Replace("Library", "Automation Studio").Replace(" - ", "\n");
            this.labelVersion.Text = "Version " + CoreWrapper.Instance.productVersion;
        }


        private void linkLabelCodecentrix_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = CatStudioConstants.CODECENTRIX_HOME_URL;
            process.StartInfo.UseShellExecute = true;
            process.Start();
        }

        private void linkLabelMail_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = CatStudioConstants.TWEBST_EMAIL_URL;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            catch (Exception ex)
            {
               MessageBox.Show(this, ex.Message, CatStudioConstants.TWEBST_PRODUCT_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
