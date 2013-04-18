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
using System.Text;



namespace CatStudio
{
    enum RawStatementType
    {
        START_UP,
        FIND_BROWSER,
        CLICK,
        TEXT_CHANGE,
        SELECTION_CHANGE,
        BACK_NAVIGATION,
        FORWARD_NAVIGATION,
        RIGHT_CLICK,
    }


    internal class RawStatement
    {
        #region Public Area

        static public RawStatement CreateStartBrowserStatement(String url, int brwsNameIndex, int brwsHwnd)
        {
            RawStatement result = new RawStatement();

            result.type           = RawStatementType.START_UP;
            result.browserURL     = url;
            result.browserNameIdx = brwsNameIndex;
            result.browserHWND    = brwsHwnd;

            return result;
        }


        static public RawStatement CreateFindBrowserStatement
            (
                String url,
                String title,
                String appName,
                int brwsNameIndex,
                int brwsHwnd
            )
        {
            RawStatement result = new RawStatement();

            result.type          = RawStatementType.FIND_BROWSER;
            result.browserURL     = url;
            result.browserNameIdx = brwsNameIndex;
            result.browserHWND    = brwsHwnd;
            result.browserAppName = appName;
            result.browserTitle   = title;

            return result;
        }


        static public RawStatement CreateClickStatement(String tag, String attr, String attrVal, int index, bool isChecked, bool isRightClick, int brwsNameIndex)
        {
            RawStatement result = new RawStatement();

            result.type           = isRightClick ? RawStatementType.RIGHT_CLICK : RawStatementType.CLICK;
            result.tagName        = tag;
            result.attributeName  = attr;
            result.attributeValue = attrVal;
            result.index          = index;
            result.isChecked      = isChecked;

            return result;
        }


        static public RawStatement CreateTextChangeStatement(String tag, String attr, String attrVal, List<String> val, int index, int brwsNameIndex)
        {
            return CreateChangeStatement(RawStatementType.TEXT_CHANGE, tag, attr, attrVal, val, index, brwsNameIndex);
        }


        static public RawStatement CreateSelectionChangeStatement
            (
                String tag, String attr, String attrVal, List<String> val, bool isMultipleSel,
                bool genVarForSelect, int index, int brwsNameIndex
            )
        {
            RawStatement rs = CreateChangeStatement(RawStatementType.SELECTION_CHANGE, tag, attr, attrVal, val, index, brwsNameIndex);

            rs.generateVarForSelect = genVarForSelect && isMultipleSel && (val.Count > 1);
            rs.isMultipleSelection  = isMultipleSel;
            return rs;
        }


        static public RawStatement CreateBackStatement()
        {
            RawStatement backStatement = new RawStatement();
            backStatement.type = RawStatementType.BACK_NAVIGATION;

            return backStatement;
        }


        static public RawStatement CreateForwardStatement()
        {
            RawStatement fwdStatement = new RawStatement();
            fwdStatement.type = RawStatementType.FORWARD_NAVIGATION;

            return fwdStatement;
        }


        public override bool Equals(Object obj)
        {
            if (obj == null || (GetType() != obj.GetType()))
            {
                return false;
            }

            RawStatement otherRawStat = (RawStatement)obj;

            return ((this.isMultipleSelection == otherRawStat.isMultipleSelection) &&
                    (this.tagName             == otherRawStat.tagName)             &&
                    (this.attributeValue      == otherRawStat.attributeValue)      &&
                    (this.attributeName       == otherRawStat.attributeName)       &&
                    (this.type                == otherRawStat.type)                &&
                    (this.browserURL          == otherRawStat.browserURL));
        }


        public override int GetHashCode()
        {
            // To avoid warning.
            return base.GetHashCode();
        }


        public bool IsChecked
        {
            get { return this.isChecked; }
        }


        public int Index
        {
            get { return this.index; }
        }


        public RawStatementType Type
        {
            get { return this.type; }
        }


        public String TagName
        {
            get { return this.tagName; }
        }


        public String Url
        {
            get { return this.browserURL; }
        }


        public String AttrName
        {
            get { return this.attributeName; }
        }

        
        public String AttrValue
        {
            get { return this.attributeValue; }
        }


        public bool IsMultipleSelection
        {
            get { return this.isMultipleSelection; }
        }


        public List<String> Values
        {
            get { return this.values; }
        }


        public bool DeclareSelectVar
        {
            get { return this.generateVarForSelect;  }
            set { this.generateVarForSelect = value; }
        }

        
        public int BrowserNameIdx
        {
            get { return this.browserNameIdx; }
        }


        public int BrowserHWND
        {
            get { return this.browserHWND; }
        }

        public String BrowserTitle
        {
            get { return this.browserTitle; }
        }

        public String BrowserAppName
        {
            get { return this.browserTitle; }
        }

        #endregion


        #region Private Area

        static private RawStatement CreateChangeStatement
            (
                RawStatementType t, String tag, String attr,
                String attrVal, List<String> val, int index, int brwsNameIndex
            )
        {
            RawStatement result = new RawStatement();

            result.type           = t;
            result.tagName        = tag;
            result.attributeName  = attr;
            result.attributeValue = attrVal;
            result.values         = val;
            result.index          = index;

            return result;
        }


        private RawStatement()
        {
        }


        private RawStatementType type;
        private String           tagName;
        private String           browserURL;
        private String           browserTitle;
        private String           browserAppName;
        private List<String>     values;
        private String           attributeName;
        private String           attributeValue;
        private bool             generateVarForSelect = false;
        private bool             isMultipleSelection  = false;
        private int              index                = 0;
        private bool             isChecked            = false;
        private int              browserNameIdx       = 0;
        private int              browserHWND          = -1;

        #endregion
    }
}
