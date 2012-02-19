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
using System.Windows.Forms;
using OpenTwebstLib;
using System.Threading;



namespace CatStudio
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
			// Allow one instance only.
			bool  isFirstInstance = false;
			Mutex mutex           = new Mutex(true, "opentwebst{8008D7FE-0420-4a30-A2FA-AF0057865F6E}", out isFirstInstance);
			if (!isFirstInstance)
			{
				return;
			}

            using (mutex)
            {
                // Try to create Twebst core object to catch the case when OpenTwebstLib.dll is not properly registered.
                try
                {
                    ICore core = CoreWrapper.Instance;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    DialogResult res = MessageBox.Show("It seems that \"Open Twebst\" is not properly installed!\nPlease re-install the product.",
                                                       CatStudioConstants.TWEBST_PRODUCT_NAME,
                                                       MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new BrowserForm());
            }
        }
    }
}
