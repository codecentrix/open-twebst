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
    class WatirGenerator : BaseLanguageGenerator, ICustomRecording
    {
        public WatirGenerator()
        {
            InitWatirTagDict();

            this.CLICK_ACTION_CALL   = "click";
            this.CHECK_ACTION_CALL   = "set";
            this.UNCHECK_ACTION_CALL = "clear";

            this.START_UP_STATEMENT                         = "require 'rubygems'\nrequire 'watir'\n\nbrowser = Watir::Browser.new\nbrowser.goto '{0}'\n";

            this.CLICK_NO_INDEX_STATEMENT                   = "browser.{0}(:{1} => '{2}').{3}";
            this.CLICK_NO_INDEX_NO_ATTR_STATEMENT           = "browser.{0}().{1}";
            this.CLICK_STATEMENT                            = "browser.{0}(:{1} => '{2}', :index => {3}).{4}";
            this.CLICK_NO_ATTR_STATEMENT                    = "browser.{0}(:index => {1}).{2}";

            this.TEXT_CHANGED_NO_INDEX_STATEMENT            = "browser.{0}(:{1} => '{2}').set '{3}'";
            this.TEXT_CHANGED_NO_INDEX_NO_ATTR_STATEMENT    = "browser{0}().set '{1}'";
            this.TEXT_CHANGED_STATEMENT                     = "browser{0}(:{1} => '{2}', :index => {3}).set '{4}'";
            this.TEXT_CHANGED_NO_ATTR_STATEMENT             = "browser{0}(:index => {1}).set '{2}'";
            this.TEXT_CHANGED_ON_FILE_IE8_COMMENT           = "# Because of new HTML 5 security specifications, IE8 - IE9 does not reveal the real local path of the file you have selected. You have to manually change the code";

            this.SELECT_MULTIPLE_DECLARATION                = "";
            this.SELECT_MULTIPLE_FIRST_ITEM_STATEMENT       = "s.select '{0}'";
            this.SELECT_MULTIPLE_NO_INDEX_STATEMENT         = "{0}s = browser.{1}(:{2} => '{3}')\n";
            this.SELECT_MULTIPLE_NO_INDEX_NO_ATTR_STATEMENT = "{0}s = browser.{1}()\n";
            this.SELECT_MULTIPLE_STATEMENT                  = "{0}s = browser.{1}(:{2} => '{3}', :index => {4})\n";
            this.SELECT_MULTIPLE_NO_ATTR_STATEMENT          = "{0}s = browser.{1}(:index => {2})\n";
            this.ADD_SELECTION_STATEMENT                    = "\ns.select '{0}'";

            this.SELECT_NO_INDEX_STATEMENT                  = "browser.{0}(:{1} => '{2}').select '{3}'";
            this.SELECT_NO_INDEX_NO_ATTR_STATEMENT          = "browser.{0}().select '{1}'";
            this.SELECT_STATEMENT                           = "browser.{0}(:{1} => '{2}', :index => {3}).select '{4}'";
            this.SELECT_NO_ATTR_STATEMENT                   = "browser.{0}(:index => {1}).select '{2}'";

            this.BACK_NAVIGATION_STATEMENT                  = "";
            this.FORWARD_NAVIGATION_STATEMENT               = "";

            this.START_UP_STATEMENT = "# The current version of Twebst Web Recorder does NOT support recording in frames/iframes.\n" +
                                      "# You have to manually add browser.frame statements.\n" +
                                      "# See: http://wiki.openqa.org/display/WTR/Frames \n\n" + 
                                      this.START_UP_STATEMENT;
        }


        protected override String EscapeStr(String source)
        {
            if (source == null)
            {
                return null;
            }

            String result = source.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n");
            return result;
        }


        internal override void Play(String code)
        {
            PlayFile(code, FileExt);
        }


        internal override String FileExt
        {
            get { return ".rb"; }
        }


        public override string ToString()
        {
            return "Watir";
        }

        internal override Encoding GeneratorEncoding
        {
            // Python code is ASCII encoded.
            get { return Encoding.ASCII; }
        }

        protected override String EncodeHtmlTagName(String tagName)
        {
            String watirTag;
            if (watirTags.TryGetValue(tagName, out watirTag))
            {
                return watirTag;
            }
            else
            {
                return tagName;
            }
        }


        protected override String EncodeAttrName(String attrName)
        {
            if ((attrName == "innertext") || (attrName == "uiname"))
            {
                return "text";
            }
            else
            {
                return attrName;
            }
        }


        protected override int EncodeIndex(int index)
        {
            // Watir indexes are 1-based (and not zero-based like in Twebst).
            return index + 1;
        }


        private void InitWatirTagDict()
        {
            watirTags["a"]              = "link";
            watirTags["input text"]     = "text_field";
            watirTags["input password"] = "text_field";
            watirTags["input button"]   = "button";
            watirTags["input submit"]   = "button";
            watirTags["input reset" ]   = "button";
            watirTags["input image"]    = "button";
            watirTags["input checkbox"] = "checkbox";
            watirTags["input file"]     = "file_field";
            watirTags["input radio"]    = "radio";
            watirTags["td"]             = "cell";
            watirTags["tr"]             = "row";
            watirTags["img"]            = "image";
            watirTags["select"]         = "select_list";
            watirTags["innertext"]      = "text";
        }

        private Dictionary<String, String> watirTags = new Dictionary<String,String>();

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // ICustomRecording
        public Object[] GetRecordedAttributes(String tagName, String inputType)
        {
            bool bDonUseValue = ((inputType == "text") || (inputType == "password") || (tagName == "textarea"));
            if (bDonUseValue)
            {
                Object[] res = { "id", "name", "value", "class", "title" };
                return res;
            }
            else
            {
                Object[] res = { "id", "name", "class", "title", "alt", "for", "href", "src", "action", "innertext" };
                return res;
            }
        }

        // Don't use short SRC (stripped down to file name as in Twebst).
        public bool ShortSrc      { get { return false; } }
        public bool IncludeFrames { get { return true;  } }
        public bool LocalIndex    { get { return true;  } }
    }
}
