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
using mshtml;
using System.Collections.Generic;
using System.Text;
using OpenTwebstLib;



namespace CatStudio
{
    class RecEventArgs : EventArgs
    {
        public static RecEventArgs CreateRecEvent(IHTMLElement htmlElem, IBrowser twbstBrowser)
        {
            htmlElem = FilterHtmlElement(htmlElem);

            RecEventArgs result = new RecEventArgs();
            String       topURL = twbstBrowser.url;

            result.htmlTargetElem = htmlElem;
            result.browser        = twbstBrowser;
            result.tagName        = htmlElem.tagName.ToLower();
            result.browserURL     = topURL;
            result.browserTitle   = twbstBrowser.title;
            result.browserAppName = twbstBrowser.app;
            result.browserHwnd    = 0; /*twbstBrowser.nativeBrowser.HWND;*/ // TODO: gotta solve exception? Maybe we don't need the handle after all...
            result.BuildAttributeDictionary(htmlElem);
            result.BuildValuesList(htmlElem);

            IHTMLInputElement inputElem = htmlElem as IHTMLInputElement;
            if (inputElem != null)
            {
                result.inputType = inputElem.type.ToLower();

                if ((result.inputType == "checkbox") || (result.inputType == "radio"))
                {
                    result.isChecked = inputElem.@checked;
                }
            }

            IHTMLSelectElement selectElem = htmlElem as IHTMLSelectElement;
            if (selectElem != null)
            {
                result.isMultipleSelection = selectElem.multiple;
            }

            return result;
        }

        public IHTMLElement TargetElement
        {
            get { return this.htmlTargetElem; }
        }


        public IBrowser Browser
        {
            get { return this.browser; }
        }


        public List<String> Values
        {
            get { return this.values; }
        }


        public String TagName
        {
            get { return this.tagName; }
        }


        public String InputType
        {
            get { return this.inputType; }
        }


        public String BrowserURL
        {
            get { return this.browserURL; }
        }


        public int BrowserHWND
        {
            get { return this.browserHwnd; }
        }

        public String BrowserTitle
        {
            get { return this.browserTitle; }
        }

        public String BrowserAppName
        {
            get { return this.browserAppName; }
        }

        public bool IsMultipleSelection
        {
            get { return this.isMultipleSelection; }
        }


        public bool IsChecked
        {
            get { return this.isChecked; }
        }


        public String GetAttribute(String attrName)
        {
            String attrValue;
            
            attrName = attrName.ToLower();

            if (this.attributeMap.TryGetValue(attrName, out attrValue))
            {
                return attrValue;
            }
            else
            {
                return "";
            }
        }


        public bool GetFirstAttribute(out String attrName, out String attrVal)
        {
            foreach (KeyValuePair<string, string> kvp in this.attributeMap)
            {
                attrName = kvp.Key;
                attrVal  = kvp.Value;

                return true;
            }

            attrName = null;
            attrVal  = null;

            return false;
        }


        private RecEventArgs()
        {
        }


        private void BuildValuesList(IHTMLElement htmlElem)
        {
            String tagName = htmlElem.tagName.ToLower();

            if ("input" == tagName)
            {
                String val = ((IHTMLInputElement)htmlElem).value;
                if (String.IsNullOrEmpty(val))
                {
                    String type = ((IHTMLInputElement)htmlElem).type;
                    if ((type != null) && type.ToLower() == "file")
                    {
                        val = "insert file path here";   
                    }
                    else
                    {
                        val = "insert text here";
                    }
                }

                this.values.Add(val);
            }
            else if ("select" == tagName)
            {
                IHTMLSelectElement selectElem = (IHTMLSelectElement)htmlElem;
                int    len        = selectElem.length;
                Object dummyIndex = 0;

                for (int i = 0; i < len; ++i)
                {
                    Object       crntIndex  = i;
                    IHTMLElement crntElem = (IHTMLElement)selectElem.item(crntIndex, dummyIndex);

                    IHTMLOptionElement crntOption = crntElem as IHTMLOptionElement;
                    if (crntOption != null)
                    {
                        if (crntOption.selected)
                        {
                            // I don't know why spaces appears like 0xA0 (no break space).
                            String crntOptionText = crntOption.text;
                            if (crntOptionText != null)
                            {
                                crntOptionText = crntOptionText.Replace('\xA0', '\x20');
                                this.values.Add(crntOptionText);
                            }
                            else
                            {
                                IHTMLOptionElement3 crntOption3 = crntOption as IHTMLOptionElement3;
                                if (crntOption3 != null)
                                {
                                    crntOptionText = crntOption3.label;
                                    if (crntOptionText != null)
                                    {
                                        crntOptionText = crntOptionText.Replace('\xA0', '\x20');
                                        this.values.Add(crntOptionText);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if ("textarea" == tagName)
            {
                String val = ((IHTMLInputElement)htmlElem).value;
                if (String.IsNullOrEmpty(val))
                {
                    val = "insert your text here";
                }

                this.values.Add(val);
            }
        }


        private void BuildAttributeDictionary(IHTMLElement htmlElem)
        {
            IHTMLDOMNode             htmlNode       = (IHTMLDOMNode)htmlElem;
            IHTMLAttributeCollection attrCollection = (IHTMLAttributeCollection)htmlNode.attributes;

            for (int i = 0; i < attrCollection.length; ++i)
            {
                Object crntIndex = i;
                IHTMLDOMAttribute crntAttribute = (IHTMLDOMAttribute)attrCollection.item(ref crntIndex);

                String nodeName  = ((String)(crntAttribute.nodeName)).ToLower();
                if (nodeName != CatStudioConstants.HOOKED_BY_REC_ATTR)
                {
                    if ((nodeName == "src")   || (nodeName == "href"  ) ||
                        (nodeName == "id")    || (nodeName == "name"  ) ||
                        (nodeName == "class") || (nodeName == "alt"   ) ||
                        (nodeName == "title") || (nodeName == "action") ||
                        (nodeName == "for")   || (nodeName == "value"))
                    {
                        String nodeValue = crntAttribute.nodeValue as String;
                        if (nodeValue != null)
                        {
                            this.attributeMap.Add(nodeName, nodeValue);
                        }
                    }
                }
            }

            // Add "uiName" pseudo-attribute to dictionary.
            IElement twbstElem = this.browser.core.AttachToNativeElement(htmlElem);
            String   textAttr  = twbstElem.uiName.Trim(); // Remove blanks from start/end of the text.

            // Skip too long texts or empty strings.
            if (!String.IsNullOrEmpty(textAttr) && (textAttr.Length <= CatStudioConstants.MAX_TEXT_ATTR_LEN_TO_RECORD))
            {
                this.attributeMap.Add("uiname", textAttr);
            }

            // Add innerText for Watir recorder.
            String innerText = htmlElem.innerText;
            if (!String.IsNullOrEmpty(innerText) && (innerText.Length <= CatStudioConstants.MAX_TEXT_ATTR_LEN_TO_RECORD))
            {
                this.attributeMap.Add("innertext", innerText);
            }
        }


        private static IHTMLElement FilterHtmlElement(IHTMLElement htmlElem)
        {
            String[] tags = { "b", "i", "u", "font", "em", "cite", "mark", "strong", "small", "sub", "sup", "q", "s", "kbd", "ins" };
            String   tag  = htmlElem.tagName.ToLower();

            if (Array.IndexOf(tags, tag) >= 0)
            {
                // Find a good parent.
                String[]     parentTags = { "a", "label", "li", "ol", "ul", "dd", "dt" };
                IHTMLElement crntElem = htmlElem;

                while (true)
                {
                    IHTMLElement parentElem = crntElem.parentElement;
                    if (parentElem == null)
                    {
                        return htmlElem;
                    }

                    String parentTag = parentElem.tagName.ToLower();
                    if (Array.IndexOf(parentTags, parentTag) >= 0)
                    {
                        return parentElem;
                    }

                    crntElem = parentElem;
                }
            }

            return htmlElem;
        }


        private IHTMLElement               htmlTargetElem     = null;
        private IBrowser                   browser            = null;
        private String                     tagName;
        private String                     inputType;
        private String                     browserURL;
        private String                     browserTitle;
        private String                     browserAppName;
        private int                        browserHwnd         = 0;
        private bool                       isMultipleSelection = false;
        private bool                       isChecked           = false;
        private List<String>               values              = new List<String>();
        private Dictionary<String, String> attributeMap        = new Dictionary<String, String>();
    }
}
