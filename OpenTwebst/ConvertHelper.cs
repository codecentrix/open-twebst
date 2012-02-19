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
using mshtml;
using SHDocVw;
using Accessibility;
using System.Runtime.InteropServices;



namespace CatStudio
{
    class ConvertHelper
    {
        public static IWebBrowser2 BrowserFromHWND(IntPtr hIeWnd)
        {
            Object accObject = null;
            AccessibleObjectFromWindow(hIeWnd, OBJID_WINDOW, ref IID_IAccessible, ref accObject);

            IAccessible accessible = accObject as IAccessible;
            if (accessible != null)
            {
                IServiceProvider serviceProvider = accessible as IServiceProvider;
                if (serviceProvider != null)
                {
                    Object wndObject = null;
                    serviceProvider.QueryService(ref IID_IHTMLWindow2, ref IID_IHTMLWindow2, out wndObject);

                    IHTMLWindow2 htmlWindow = wndObject as IHTMLWindow2;
                    if (htmlWindow != null)
                    {
                        IHTMLWindow2     topHtmlWnd            = htmlWindow.top;
                        IServiceProvider windowServiceProvider = topHtmlWnd as IServiceProvider;

                        if (windowServiceProvider != null)
                        {
                            Object browserObject = null;
                            windowServiceProvider.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out browserObject);

                            IWebBrowser2 htmlBrowser = browserObject as IWebBrowser2;
                            return htmlBrowser;
                        }
                    }
                }
            }

            return null;
        }

        // This is the COM IServiceProvider interface, not System.IServiceProvider .Net interface!
        [ComImport(), ComVisible(true), Guid("6D5140C1-7436-11CE-8034-00AA006009FA"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IServiceProvider
        {
	        [return: MarshalAs(UnmanagedType.I4)][PreserveSig]
	        int QueryService(ref Guid guidService, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
        }

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow
            (IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        private const int OBJID_WINDOW = 0;
        private static Guid IID_IHTMLWindow2   = new Guid("{332c4427-26cb-11d0-b483-00c04fd90119}");
        private static Guid IID_IAccessible    = new Guid("{618736e0-3c3d-11cf-810c-00aa00389b71}");
        private static Guid IID_IWebBrowser2   = new Guid("{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}");
        private static Guid IID_IWebBrowserApp = new Guid("{0002DF05-0000-0000-C000-000000000046}");
    }
}
