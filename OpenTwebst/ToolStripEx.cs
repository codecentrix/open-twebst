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
using System.Windows.Forms;
using System.Text;


namespace CatStudio
{
    class ToolStripEx : ToolStrip
    {
        // Gets or sets whether the ToolStripEx accepts item clicks when its containing form does not have input focus.
        // Default value is false, which is the same behavior provided by the base ToolStrip class.
        public bool ClickThrough
        {
            get
            {
                return this.clickThrough;
            }
            set
            {
                this.clickThrough = value;
            }
        }



        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (this.clickThrough && (m.Msg == Win32Api.WM_MOUSEACTIVATE) &&
                (m.Result == (IntPtr)Win32Api.MA_ACTIVATEANDEAT))
            {
                m.Result = (IntPtr)Win32Api.MA_ACTIVATE;
            }
        }


        private bool clickThrough = false;
    }
}
