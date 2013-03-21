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
using System.Drawing;


namespace CatStudio
{
    class Win32HookMsgEventArgs : EventArgs
    {
        public enum HookMsgType
        {
            KEY_PRESSED_MSG,
            CLICK_DOWN_MSG,
            CLICK_UP_MSG,
            RIGHT_CLICK_DOWN_MSG,
            RIGHT_CLICK_UP_MSG,
        }

        public Win32HookMsgEventArgs(IntPtr hIEWnd, HookMsgType msgType)
        {
            this.hWnd = hIEWnd;
            this.type = msgType;
        }


        public IntPtr WndHandle
        {
            get { return this.hWnd; }
        }


        public HookMsgType Type
        {
            get { return this.type; }
        }

        private IntPtr      hWnd;
        private HookMsgType type;
    }



    class Win32GetMsgHook
    {
        #region Public Area

        public event EventHandler<Win32HookMsgEventArgs> Win32HookMsg;

        public void Install()
        {
            if ((hHook != 0) || (getMsgHookProcedure != null))
            {
                throw new Exception("Hook already installed in Hooks.Install");
            }

            getMsgHookProcedure = new Win32Api.HookProc(GetMsgHookProc);

            hHook = Win32Api.SetWindowsHookEx(Win32Api.WH_GETMESSAGE, getMsgHookProcedure, (IntPtr)0, AppDomain.GetCurrentThreadId());
            if (hHook == 0)
            {
                throw new Exception("SetWindowsHookEx failed in Hooks.Install");
            }
        }


        public void UnInstall()
        {
            if (0 == hHook)
            {
                throw new Exception("hHook is zero in Hooks.UnInstall");
            }

            bool ret = Win32Api.UnhookWindowsHookEx(hHook);
            if (ret == false)
            {
                throw new Exception("UnhookWindowsHookEx failed in Hooks.Install");
            }

            hHook               = 0;
            getMsgHookProcedure = null;
        }

        #endregion

        
        #region Private Area
        private Win32HookMsgEventArgs.HookMsgType GetMsgType(int message)
        {
            switch (message)
            {
                case Win32Api.WM_LBUTTONUP:
                {
                    return Win32HookMsgEventArgs.HookMsgType.CLICK_UP_MSG;
                }

                case Win32Api.WM_LBUTTONDOWN:
                {
                    return Win32HookMsgEventArgs.HookMsgType.CLICK_DOWN_MSG;
                }

                case Win32Api.WM_RBUTTONDOWN:
                {
                    return Win32HookMsgEventArgs.HookMsgType.RIGHT_CLICK_DOWN_MSG;
                }

                case Win32Api.WM_RBUTTONUP:
                {
                    return Win32HookMsgEventArgs.HookMsgType.RIGHT_CLICK_UP_MSG;
                }

                case Win32Api.WM_KEYDOWN:
                {
                    return Win32HookMsgEventArgs.HookMsgType.KEY_PRESSED_MSG;
                }

                default:
                {
                    throw new Exception("Invalid message type in Hooks.GetMsgType");
                }
            }
        }


        private int GetMsgHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam);
            }
            else
            {
                GetMsgHookStruct msgStruct = (GetMsgHookStruct)Marshal.PtrToStructure(lParam, typeof(GetMsgHookStruct));

                switch (msgStruct.message)
                {
                    case Win32Api.WM_LBUTTONUP:
                    case Win32Api.WM_LBUTTONDOWN:
                    case Win32Api.WM_RBUTTONDOWN:
                    {
                        // With WH_GETMESSAGE hooks, it's very important to check for WPARAM equal to PM_REMOVE
                        // because otherwise you may end up processing the event twice. Windows calls your WH_GETMESSAGE
                        // hook whenever the app calls ::GetMessage or ::PeekMessage. Many apps call ::PeekMessage to check
                        // for messages before fetching them, in order to do idle processing between messages. When the
                        // app calls ::PeekMessage, Windows calls your hook with WPARAM set to PM_NOREMOVE.
                        if (wParam.ToInt32() == Win32Api.PM_REMOVE)
                        {
                            if (Win32Api.IsIEServerOrChild(msgStruct.hwnd))
                            {
                                //System.Diagnostics.Trace.WriteLine("OnNewWin32Msg event rised for " + msgStruct.message + " on " + msgStruct.hwnd);
                                Win32HookMsgEventArgs.HookMsgType msgType = GetMsgType(msgStruct.message);
                                OnNewWin32Msg(msgStruct.hwnd, msgType);
                            }
                        }

                        break;
                    }

                    case Win32Api.WM_KEYDOWN:
                    {
                        if (wParam.ToInt32() == Win32Api.PM_REMOVE)
                        {
                            if (Win32Api.IsIEServerOrChild(msgStruct.hwnd))
                            {
                                OnNewWin32Msg(msgStruct.hwnd, Win32HookMsgEventArgs.HookMsgType.KEY_PRESSED_MSG);
                            }
                        }

                        break;
                    }
                }

                return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam);
            }
        }


        private void OnNewWin32Msg(IntPtr hIEWnd, Win32HookMsgEventArgs.HookMsgType type)
        {
            if (Win32HookMsg != null)
            {
                Win32HookMsg(this, new Win32HookMsgEventArgs(hIEWnd, type));
            }
        }


        private Win32Api.HookProc getMsgHookProcedure = null;
        private int hHook  = 0;

        #endregion
    }


    class Win32MouseLLEventArgs : EventArgs
    {
        public Win32MouseLLEventArgs(int msg, int x, int y)
        {
            this.message = msg;
            this.screenPoint = new Point(x, y);
        }

        public Point ScreenPoint
        {
            get { return this.screenPoint; }
        }

        public int Message
        {
            get { return this.message; }
        }

        private Point screenPoint;
        private int   message;
    }


    class Win32LLMouseHook
    {
        #region Public Area

        public event EventHandler<Win32MouseLLEventArgs> Win32HookMouseMsg;

        public void Install()
        {
            if ((hHook != 0) || (mouseHookProcedure != null))
            {
                throw new Exception("Hook already installed in Win32LLMouseHook.Install");
            }

            mouseHookProcedure = new Win32Api.HookProc(MouseHookProc);
            hHook = Win32Api.SetWindowsHookEx(Win32Api.WH_MOUSE_LL, mouseHookProcedure, (IntPtr)0, 0);
            if (hHook == 0)
            {
                throw new Exception("SetWindowsHookEx failed in Win32LLMouseHook.Install");
            }
        }


        public void UnInstall()
        {
            if (0 == hHook)
            {
                // Not installed.
                return;
            }

            bool ret = Win32Api.UnhookWindowsHookEx(hHook);
            if (ret == false)
            {
                throw new Exception("UnhookWindowsHookEx failed in Win32LLMouseHook.Install");
            }

            hHook = 0;
            mouseHookProcedure = null;
        }

        #endregion

        #region Private Area
        private int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                // If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx
                // function without further processing and should return the value returned by CallNextHookEx.
                return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam);
            }
            else
            {
                if (Win32HookMouseMsg != null)
                {
                    MouseLLHookStruct msgStruct = (MouseLLHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseLLHookStruct));
                    int msg = wParam.ToInt32();
                    switch (msg)
                    {
                        case Win32Api.WM_LBUTTONDOWN:
                        case Win32Api.WM_LBUTTONUP:
                        {
                            IntPtr topWnd = Win32Api.WindowFromPoint(msgStruct.pt);
                            if (Win32Api.IsIEServerWindow(topWnd))
                            {
                                Win32HookMouseMsg(this, new Win32MouseLLEventArgs(msg, msgStruct.pt.x, msgStruct.pt.y));
                                return 1; // Prevent the message being dispatched to target window.
                            }

                            return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam); 
                        }

                        case Win32Api.WM_MOUSEMOVE:
                        {
                            IntPtr topWnd = Win32Api.WindowFromPoint(msgStruct.pt);
                            if (Win32Api.IsIEServerWindow(topWnd))
                            {
                                Win32HookMouseMsg(this, new Win32MouseLLEventArgs(msg, msgStruct.pt.x, msgStruct.pt.y));
                            }

                            return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam);
                        }
                    }
                }

                return Win32Api.CallNextHookEx(hHook, nCode, wParam, lParam);
            }
        }
        
        private Win32Api.HookProc mouseHookProcedure = null;
        private int hHook  = 0;

        #endregion
    }
}
