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
using OpenTwebstLib;
using mshtml;



namespace CatStudio
{
    class CodeGenEventArgs : EventArgs
    {
        public CodeGenEventArgs(String code)
        {
            this.code = code;
        }

        public String Code
        {
            get { return this.code;  }
        }

        private String code;
    }


    class CodeGenerator
    {
        #region Public Area

        public event EventHandler<CodeGenEventArgs> NewStatement;
        public event EventHandler<CodeGenEventArgs> CodeChanged; // Language changed or last statement deleted.


        public CodeGenerator(BaseLanguageGenerator lg)
        {
            this.languageGenerator = lg;

            Recorder.Instance.ClickAction      += OnClickAction;
            Recorder.Instance.RightClickAction += OnRightClickAction;
            Recorder.Instance.ChangeAction     += OnChangeAction;
            Recorder.Instance.RecStopped       += OnRecordStopped;
            Recorder.Instance.RecStarted       += OnRecordStart;

            // Better not record that! To play back/forward with C#/VB.Net is hard because Microsoft.mshtml.dll is referenced.
            //Recorder.Instance.BackAction    += OnRecordBack;
            //Recorder.Instance.ForwardAction += OnRecordForward;
        }


        public String GetAllCode()
        {
            return this.allCode;
        }


        public String GetAllDecoratedCode()
        {
            return this.languageGenerator.DecorateCode(this.allCode);
        }


        public BaseLanguageGenerator Language
        {
            get { return this.languageGenerator; }

            set
            {
                if (this.languageGenerator != value)
                {
                    this.languageGenerator = value;
                    this.FireCodeChanged();
                }
            }
        }


        public void Reset()
        {
            this.startupCodeGenerated  = false;
            this.snapshotCodeGenerated = false;
            this.allCode               = "";
            this.rawStatements.Clear();
        }


        // On first action a snapshot of all controls will be generated according to this settings.
        public bool GenerateSnapshot
        {
            get { return this.generateSnapshot; }
            set { this.generateSnapshot = value; }
        }


        public void DeleteLastStatement()
        {
            if (this.rawStatements.Count > 0)
            {
                this.rawStatements.RemoveAt(this.rawStatements.Count - 1);

                if (this.rawStatements.Count == 0)
                {
                    this.Reset();
                }

                this.FireCodeChanged();
            }
        }

        #endregion


        #region Private Area


        private bool IsSelectVarDeclared
        {
            get
            {
                foreach (RawStatement rs in this.rawStatements)
                {
                    if (rs.DeclareSelectVar)
                    {
                        return true;
                    }
                }

                return false;
            }
        }


        private int ComputeIndex(RecEventArgs evt, String twebstTagName, String attr, String attrVal)
        {
            try
            {
                String           uniqueValue     = DateTime.Now.ToString();
                IHTMLElement     htmlTarget      = evt.TargetElement;
                String           searchCondition = null;
                IElementList     elements        = null;
                bool             isInnerText     = (attr == "innertext");
                ICustomRecording customRec       = this.languageGenerator as ICustomRecording;

                htmlTarget.setAttribute(CatStudioConstants.FIND_INDEX_HELPER_ATTR, uniqueValue, 0);

                if ((attr != null) && (attrVal != null) && !isInnerText)
                {
                    searchCondition = String.Format("{0}={1}", attr, attrVal);

                    if ((customRec == null) || !customRec.LocalIndex)
                    {
                        elements = evt.Browser.FindAllElements(twebstTagName, searchCondition);
                    }
                    else
                    {
                        ICore    core       = CoreWrapper.Instance;
                        IElement targetElem = core.AttachToNativeElement(evt.TargetElement);

                        // Search only elements in current frame.
                        elements = targetElem.parentFrame.FindChildrenElements(twebstTagName, searchCondition);
                    }
                }
                else
                {
                    if ((customRec == null) || !customRec.LocalIndex)
                    {
                        elements = evt.Browser.FindAllElements(twebstTagName, "");
                    }
                    else
                    {
                        ICore    core       = CoreWrapper.Instance;
                        IElement targetElem = core.AttachToNativeElement(evt.TargetElement);

                        // Search only elements in current frame.
                        elements = targetElem.parentFrame.FindChildrenElements(twebstTagName, "");
                    }
                }

                int index = 0;
                int length = elements.length;

                if (length <= 1)
                {
                    return 0;
                }

                if (!isInnerText)
                {
                    for (int i = 0; i < length; ++i)
                    {
                        String findIndexAttrValue = (String)elements[i].GetAttribute(CatStudioConstants.FIND_INDEX_HELPER_ATTR);

                        if (uniqueValue == findIndexAttrValue)
                        {
                            index = i;
                            break;
                        }
                    }
                }
                else
                {
                    // Watir uses innerText but Twebst does not search for innerText.
                    for (int i = 0; i < length; ++i)
                    {
                        String findIndexAttrValue = (String)elements[i].GetAttribute(CatStudioConstants.FIND_INDEX_HELPER_ATTR);
                        if (uniqueValue == findIndexAttrValue)
                        {
                            break;
                        }

                        if (attrVal == elements[i].nativeElement.innerText)
                        {
                            index++;
                        }
                    }
                }

                return index;
            }
            catch
            {
                // Something went wrong, give up using index.
                return 0;
            }
        }


        private bool GetAttrToRecord(RecEventArgs evt, out String attr, out String attrVal)
        {
            ICustomRecording customRec = this.languageGenerator as ICustomRecording;
            if (customRec != null)
            {
                return this.FindNonEmptyAttribute(evt, out attr, out attrVal, customRec.GetRecordedAttributes(evt.TagName, evt.InputType));
            }

            // href is for area, src is for img.
            return this.FindNonEmptyAttribute(evt, out attr, out attrVal, "id", "name", "uiname", "src", "href", "class");
        }


        private bool FindNonEmptyAttribute(RecEventArgs evt, out String attr, out String attrVal, params Object[] attrNames)
        {
            ICustomRecording customRec = this.languageGenerator as ICustomRecording;

            for (int i = 0; i < attrNames.Length; ++i)
            {
                String crntAttr = (String)attrNames[i];
                String val      = evt.GetAttribute(crntAttr);

                if (val != "")
                {
                    attr = crntAttr;

                    if ((attr == "src") && ((customRec == null) || customRec.ShortSrc))
                    {
                        // For src attribute of the image element keep only the file name if any.
                        int lastSlashIndex = val.LastIndexOf('/');

                        attrVal = val.Substring(lastSlashIndex + 1);
                    }
                    else
                    {
                        attrVal = val;
                    }

                    return true;
                }
            }

            return evt.GetFirstAttribute(out attr, out attrVal);
        }


        private String ComputeTwebstTagName(RecEventArgs evt)
        {
            if (evt.InputType != null)
            {
                return evt.TagName + " " + evt.InputType;
            }
            else
            {
                return evt.TagName;
            }
        }

        
        private void OnRecordBack(Object sender, EventArgs evt)
        {
            if (this.startupCodeGenerated)
            {
                RawStatement backStatement = RawStatement.CreateBackStatement();
                this.AddStatement(backStatement);
            }
        }


        private void OnRecordForward(Object sender, EventArgs evt)
        {
            if (this.startupCodeGenerated)
            {
                RawStatement fwdtatement = RawStatement.CreateForwardStatement();
                this.AddStatement(fwdtatement);
            }
        }


        private void OnRecordStart(Object sender, EventArgs evt)
        {
        }


        private void OnRecordStopped(Object sender, EventArgs evt)
        {
        }


        private void OnRightClickAction(Object sender, RecEventArgs evt)
        {
            OnClick(sender, evt, true);
        }


        private void OnClickAction(Object sender, RecEventArgs evt)
        {
            OnClick(sender, evt, false);
        }


        private void OnClick(Object sender, RecEventArgs evt, bool isRightClick)
        {
            GenStartup(evt.BrowserURL);
            GenSnapshot(evt.Browser);

            String attr;
            String attrVal;

            String twebstTagName = ComputeTwebstTagName(evt);
            bool   success       = GetAttrToRecord(evt, out attr, out attrVal);

            // What to do if no attribute (success is false). Use only index or don't record anything?
            int          index = ComputeIndex(evt, twebstTagName, attr, attrVal);
            RawStatement rs    = RawStatement.CreateClickStatement(twebstTagName, attr, attrVal, index, evt.IsChecked, isRightClick);

            this.AddStatement(rs);
        }


        private void OnChangeAction(Object sender, RecEventArgs evt)
        {
            GenStartup(evt.BrowserURL);
            GenSnapshot(evt.Browser);

            String attr;
            String attrVal;

            bool   success       = GetAttrToRecord(evt, out attr, out attrVal);
            String twebstTagName = ComputeTwebstTagName(evt);
            int    index         = ComputeIndex(evt, twebstTagName, attr, attrVal);

            if (evt.TagName == "select")
            {
                RawStatement rs = RawStatement.CreateSelectionChangeStatement(twebstTagName, attr, attrVal, evt.Values, evt.IsMultipleSelection, !this.IsSelectVarDeclared, index);

                if (evt.IsMultipleSelection)
                {
                    RawStatement lastStatement = this.LastStatement;
                    if (lastStatement != null)
                    {
                        if (lastStatement.Equals(rs))
                        {
                            this.rawStatements[this.rawStatements.Count - 1] = rs;
                            this.FireCodeChanged();
                        }
                        else
                        {
                            this.AddStatement(rs);
                        }
                    }
                    else
                    {
                        this.AddStatement(rs);
                    }
                }
                else
                {
                    this.AddStatement(rs);
                }
            }
            else
            {
                RawStatement rs = RawStatement.CreateTextChangeStatement(twebstTagName, attr, attrVal, evt.Values, index);
                this.AddStatement(rs);
            }
        }


        private void GenStartup(String url)
        {
            if (!this.startupCodeGenerated)
            {
                RawStatement startup = RawStatement.CreateStartupStatement(url);

                this.AddStatement(startup);
                this.startupCodeGenerated = true;
            }
        }


        private void GenSnapshot(IBrowser browser)
        {
            if (this.generateSnapshot && !this.snapshotCodeGenerated)
            {
                // TODO: Generate snapshot code for browser.
                this.snapshotCodeGenerated = true;
            }
        }


        private void AddStatement(RawStatement rs)
        {
            this.rawStatements.Add(rs);
            this.EmitStatement(rs);
        }


        private void EmitStatement(RawStatement rs)
        {
            String newStatement = this.languageGenerator.Generate(rs);
            if (newStatement != null)
            {
                newStatement += "\n";
                this.allCode += newStatement;

                if (NewStatement != null)
                {
                    this.NewStatement(this, new CodeGenEventArgs(newStatement));
                }
            }
        }


        private RawStatement LastStatement
        {
            get
            {
                if (this.rawStatements.Count == 0)
                {
                    // The statement list is empty.
                    return null;
                }

                RawStatement lastStatement = this.rawStatements[this.rawStatements.Count - 1];
                return lastStatement;
            }
        }


        private void FireCodeChanged()
        {
            this.allCode = this.languageGenerator.Generate(this.rawStatements);
            if (CodeChanged != null)
            {
                this.CodeChanged(this, new CodeGenEventArgs(this.allCode));
            }
        }


        private String                allCode               = "";
        private bool                  startupCodeGenerated  = false;
        private bool                  snapshotCodeGenerated = false;
        private bool                  generateSnapshot      = false;
        private BaseLanguageGenerator languageGenerator     = null;
        private List<RawStatement>    rawStatements         = new List<RawStatement>();

        #endregion
    }
}
