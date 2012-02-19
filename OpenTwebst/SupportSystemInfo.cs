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
using System.Management;
using System.Runtime.InteropServices;


namespace CatStudio
{
    class SupportSystemInfo
    {
        public SupportSystemInfo()
        {
            this.productName    = CoreWrapper.Instance.productName;
            this.productVersion = CoreWrapper.Instance.productVersion;
            this.ieVersion      = CoreWrapper.Instance.IEVersion;
            this.winVersion     = GetWinVersion();
        }


        public String Info
        {
            get
            {
                String result = String.Format("{0} {1}%0aInternet Explorer {2}%0a{3}",
                                              productName, productVersion,
                                              ieVersion, winVersion);

                return "---------------------------------------------------------------------------------%0a" + result;
            }
        }

        private String GetWinVersion()
        {
            try
            {
                ManagementObjectSearcher wmiSearchObj = new ManagementObjectSearcher();
                wmiSearchObj.Query.QueryString = "SELECT * FROM Win32_OperatingSystem";

                ManagementObjectCollection wmiCollObj = wmiSearchObj.Get();
                String                     osInfo     = "";

                foreach (ManagementObject mo in wmiCollObj)
                {
                    String osArchitecture = "";
                    try
                    {
                        osArchitecture = mo.Properties["OSArchitecture"].Value.ToString();
                    }
                    catch
                    {
                        // OSArchitecture is NOT supported on WinXP. Don't fail at least!
                        osArchitecture = this.GetOSArchitecture();
                    }

                    osInfo = String.Format("{0} {1} Version {2} {3}",
                                           RemoveTradeMark(mo.Properties["Caption"].Value.ToString()), osArchitecture,
                                           mo.Properties["Version"].Value, mo.Properties["CSDVersion"].Value);

                    break;
                }

                return osInfo;
            }
            catch
            {
                return "";
            }
        }


        private String GetOSArchitecture()
        {
            if (this.Is64Bit())
            {
                return "64 bit";
            }
            else
            {
                return "32 bit";
            }
        }

        
        private String RemoveTradeMark(String source)
        {
            String tradeMark = '\u2122'.ToString();
            return source.Replace(tradeMark, "");
        }


        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]  
        [return: MarshalAs(UnmanagedType.Bool)]  
        public static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool lpSystemInfo);  

        private bool Is64Bit()  
        {  
            try
            {
                bool retVal;  

                IsWow64Process(System.Diagnostics.Process.GetCurrentProcess().Handle, out retVal);  
                return retVal;
            }
            catch
            {
                // IsWow64Process requires at least WinXP SP2.
                return false;
            }
        }  



        private String productName;
        private String productVersion;
        private String ieVersion;
        private String winVersion;
    }
}
