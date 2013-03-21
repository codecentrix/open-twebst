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
using System.Runtime.InteropServices;
using Accessibility;
using mshtml;



namespace CatStudio
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT 
    {
        public POINT(System.Drawing.Point pt)
        {
            this.x = pt.X;
            this.y = pt.Y;
        }

        public int x;
        public int y;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct GetMsgHookStruct
    {
        public IntPtr hwnd;
        public int    message;
        public IntPtr wParam;
        public IntPtr lParam;
        public UInt32 time;
        public POINT  pt;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct MouseLLHookStruct
    {
        public POINT pt;
        public UInt32 mouseData;
        public UInt32 flags;
        public UInt32 time;
        public IntPtr dwExtraInfo;
    }


    class Win32Api
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet=CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", ExactSpelling=true, CharSet=CharSet.Auto)]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll")]
        public static extern IntPtr ChildWindowFromPoint(IntPtr hWndParent, POINT Point);

        [DllImport("oleacc.dll")]
        private static extern IntPtr AccessibleObjectFromPoint(POINT pt, [Out, MarshalAs(UnmanagedType.Interface)] out IAccessible accObj, [Out] out object ChildID);

        [ComImport(), ComVisible(true), Guid("6D5140C1-7436-11CE-8034-00AA006009FA"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IServiceProvider
        {
	        [return: MarshalAs(UnmanagedType.I4)][PreserveSig]
	        int QueryService(ref Guid guidService, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
        }

        public static int MakeLong (short lowPart, short highPart)
        {
            return (int)(((ushort)lowPart) | (uint)(highPart << 16));
        }

        public static short HiWord(int dword)
        {
            return (short) (dword >> 16);
        }

        public static short LoWord(int dword)
        {
            return (short) dword;
        }

        public static IAccessible GetAccessibleObjectFromPoint(System.Drawing.Point screenPoint)
        {
            try
            {
                POINT       pt         = new POINT(screenPoint);
                object      varChildID = null;
                IAccessible accObj     = null;

                AccessibleObjectFromPoint(pt, out accObj, out varChildID);
                return accObj;
            }
            catch
            {
                return null;
            }
        }

        public static IHTMLElement GetHtmlElementFromAccessible(IAccessible accObj)
        {
            IServiceProvider servProv = accObj as IServiceProvider;
            if (servProv != null)
            {
                Object htmlObj = null;
                servProv.QueryService(ref IID_IHTMLElement, ref IID_IHTMLElement, out htmlObj);

                IHTMLElement html = htmlObj as IHTMLElement;
                return html;
            }
            else
            {
                return null;
            }
        }

        public static bool IsIEServerWindow(IntPtr hWnd)
        {
            StringBuilder className = new StringBuilder(100);
            int           res       = Win32Api.GetClassName(hWnd, className, className.Capacity);
            if (res != 0)
            {
                return ("Internet Explorer_Server" == className.ToString());
            }
            else
            {
                return false;
            }
        }

        public static bool IsIEServerOrChild(IntPtr hWnd)
        {
            // On IE6 combo-boxes are implemented as child windows.
            IntPtr hCrntWnd = hWnd;

            while (true)
            {
                if (!Win32Api.IsWindow(hCrntWnd))
                {
                    return false;
                }

                if (IsIEServerWindow(hCrntWnd))
                {
                    return true;
                }

                hCrntWnd = Win32Api.GetParent(hCrntWnd);
            }
        }

        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public const int WM_KEYDOWN            = 0x0100;
        public const int WM_MOUSEMOVE          = 0x0200;
        public const int WM_LBUTTONUP          = 0x0202;
        public const int WM_LBUTTONDOWN        = 0x0201;
        public const int WM_RBUTTONDOWN        = 0x0204;
        public const int WM_RBUTTONUP          = 0x0205;
        public const int WH_GETMESSAGE         = 3;
        public const int WH_MOUSE_LL           = 14;
        public const int PM_REMOVE             = 0x0001;
        public const int WM_MOUSEACTIVATE      = 0x21;
        public const int MA_ACTIVATE           = 1;
        public const int MA_ACTIVATEANDEAT     = 2;
        public const int MA_NOACTIVATE         = 3;
        public const int MA_NOACTIVATEANDEAT   = 4;
        public static Guid IID_IHTMLElement    = new Guid("{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}");
    }


    [Guid("30510480-98b5-11cf-bb82-00aa00bdce0b")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IHTMLStyle6
    {
        [DispId(-2147412890)]
        String outline { get; set; }
    }

    [Guid("30510481-98b5-11cf-bb82-00aa00bdce0b")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IHTMLCurrentStyle5
    {
        [DispId(-2147412890)]
        String outline { get; }
    }
}
