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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Threading;



namespace CatStudio
{
    internal interface ICustomRecording
    {
        Object[] GetRecordedAttributes(String tagName, String inputType);

        bool ShortSrc      { get; }
        bool IncludeFrames { get; }
        bool LocalIndex    { get; }
    }

    public abstract class BaseLanguageGenerator
    {
        internal String Generate(RawStatement rs)
        {
            String result = null;

            switch (rs.Type)
            {
                case RawStatementType.BACK_NAVIGATION:
                {
                    result = BACK_NAVIGATION_STATEMENT;
                    break;
                }

                case RawStatementType.FORWARD_NAVIGATION:
                {
                    result = FORWARD_NAVIGATION_STATEMENT;
                    break;
                }

                case RawStatementType.START_UP:
                {
                    String startup = START_UP_STATEMENT;
                    result = String.Format(startup, rs.Url);
                    break;
                }

                case RawStatementType.RIGHT_CLICK:
                case RawStatementType.CLICK:
                {
                    String action = ((RawStatementType.CLICK == rs.Type) ? this.CLICK_ACTION_CALL : this.RIGHT_CLICK_ACTION_CALL);
                    if (rs.TagName == "input radio")
                    {
                        // For radios alwasy generate Check. To un-check a radio just click on another radio in the group.
                        action = this.CHECK_ACTION_CALL;
                    }
                    else if (rs.TagName == "input checkbox")
                    {
                        if (rs.IsChecked)
                        {
                            action = this.UNCHECK_ACTION_CALL;
                        }
                        else
                        {
                            action = this.CHECK_ACTION_CALL;
                        }
                    }

                    if (rs.Index == 0)
                    {
                        if ((rs.AttrValue != null) && (rs.AttrName != null))
                        {
                            String click   = CLICK_NO_INDEX_STATEMENT;
                            String attrVal = EscapeStr(rs.AttrValue);

                            result = String.Format(click, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, action);
                        }
                        else
                        {
                            // No relevant attribute, use only index which is zero so don't mention it.
                            String click  = CLICK_NO_INDEX_NO_ATTR_STATEMENT;
                            result = String.Format(click, EncodeHtmlTagName(rs.TagName), action);
                        }
                    }
                    else
                    {
                        if ((rs.AttrValue != null) && (rs.AttrName != null))
                        {
                            String click   = CLICK_STATEMENT;
                            String attrVal = EscapeStr(rs.AttrValue);

                            result = String.Format(click, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EncodeIndex(rs.Index), action);
                        }
                        else
                        {
                            // No relevant attribute, use only index.
                            String click   = CLICK_NO_ATTR_STATEMENT;
                            String attrVal = EscapeStr(rs.AttrValue);

                            result = String.Format(click, EncodeHtmlTagName(rs.TagName), EncodeIndex(rs.Index), action);
                        }
                    }

                    break;
                }

                case RawStatementType.TEXT_CHANGE:
                {
                    if (rs.Index == 0)
                    {
                        if ((rs.AttrValue != null) && (rs.AttrName != null))
                        {
                            String textChange = TEXT_CHANGED_NO_INDEX_STATEMENT;
                            String attrVal    = EscapeStr(rs.AttrValue);

                            result = String.Format(textChange, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EscapeStr(rs.Values[0]));
                        }
                        else
                        {
                            String textChange = TEXT_CHANGED_NO_INDEX_NO_ATTR_STATEMENT;
                            result = String.Format(textChange, EncodeHtmlTagName(rs.TagName), EscapeStr(rs.Values[0]));
                        }
                    }
                    else
                    {
                        if ((rs.AttrValue != null) && (rs.AttrName != null))
                        {
                            String textChange = TEXT_CHANGED_STATEMENT;
                            String attrVal    = EscapeStr(rs.AttrValue);

                            result = String.Format(textChange, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EncodeIndex(rs.Index), EscapeStr(rs.Values[0]));
                        }
                        else
                        {
                            String textChange = TEXT_CHANGED_NO_ATTR_STATEMENT;
                            String attrVal    = EscapeStr(rs.AttrValue);

                            result = String.Format(textChange, EncodeHtmlTagName(rs.TagName), EncodeIndex(rs.Index), EscapeStr(rs.Values[0]));
                        }
                    }

                    if (rs.TagName == "input file" )
                    {
                        if (this.IEVersion.CompareTo("8") >= 0)
                        {
                            result = TEXT_CHANGED_ON_FILE_IE8_COMMENT + "\n" + result;
                        }
                    }

                    break;
                }

                case RawStatementType.SELECTION_CHANGE:
                {
                    if (rs.IsMultipleSelection && (rs.Values.Count > 1))
                    {
                        String varDecl       = (rs.DeclareSelectVar ? SELECT_MULTIPLE_DECLARATION : "");
                        String attrVal       = EscapeStr(rs.AttrValue);
                        String textSelChange = null;
                        String selectAction  = String.Format(SELECT_MULTIPLE_FIRST_ITEM_STATEMENT, EscapeStr(rs.Values[0]));

                        if (rs.Index == 0)
                        {
                            if ((rs.AttrValue != null) && (rs.AttrName != null))
                            {
                                textSelChange = String.Format(SELECT_MULTIPLE_NO_INDEX_STATEMENT, varDecl, rs.TagName, EncodeAttrName(rs.AttrName), attrVal);
                            }
                            else
                            {
                                textSelChange = String.Format(SELECT_MULTIPLE_NO_INDEX_NO_ATTR_STATEMENT, varDecl, rs.TagName);
                            }
                        }
                        else
                        {
                            if ((rs.AttrValue != null) && (rs.AttrName != null))
                            {
                                textSelChange = String.Format(SELECT_MULTIPLE_STATEMENT, varDecl, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EncodeIndex(rs.Index));
                            }
                            else
                            {
                                textSelChange = String.Format(SELECT_MULTIPLE_NO_ATTR_STATEMENT, varDecl, EncodeHtmlTagName(rs.TagName), EncodeIndex(rs.Index));
                            }
                        }

                        result = textSelChange + selectAction;

                        // AddSelection for the rest of the items to be selected.
                        for (int i = 1; i < rs.Values.Count; ++i)
                        {
                            selectAction = String.Format(ADD_SELECTION_STATEMENT, EscapeStr(rs.Values[i]));
                            result += selectAction;
                        }

                        result += "\n";
                    }
                    else
                    {
                        if (rs.Index == 0)
                        {
                            if ((rs.AttrValue != null) && (rs.AttrName != null))
                            {
                                String attrVal      = EscapeStr(rs.AttrValue);
                                String selectChange = SELECT_NO_INDEX_STATEMENT;

                                result = String.Format(selectChange, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EscapeStr(rs.Values[0]));
                            }
                            else
                            {
                                String selectChange = SELECT_NO_INDEX_NO_ATTR_STATEMENT;
                                result = String.Format(selectChange, EncodeHtmlTagName(rs.TagName), EscapeStr(rs.Values[0]));
                            }
                        }
                        else
                        {
                            if ((rs.AttrValue != null) && (rs.AttrName != null))
                            {
                                String attrVal      = EscapeStr(rs.AttrValue);
                                String selectChange = SELECT_STATEMENT;

                                result = String.Format(selectChange, EncodeHtmlTagName(rs.TagName), EncodeAttrName(rs.AttrName), attrVal, EncodeIndex(rs.Index), EscapeStr(rs.Values[0]));
                            }
                            else
                            {
                                String selectChange = SELECT_NO_ATTR_STATEMENT;
                                result = String.Format(selectChange, EncodeHtmlTagName(rs.TagName), EncodeIndex(rs.Index), EscapeStr(rs.Values[0]));
                            }
                        }
                    }

                    break;
                }
            }

            return result;
        }


        internal String Generate(List<RawStatement> lrs)
        {
            String result = "";

            foreach (RawStatement rs in lrs)
            {
                result += Generate(rs) + "\n";
            }

            return result;
        }

        
        internal virtual String DecorateCode(String code)
        {
            // By default don't decorate.
            return code;
        }


        internal abstract void Play(String code);


        internal abstract String FileExt
        {
            get;
        }


        internal virtual Encoding GeneratorEncoding
        {
            get { return Encoding.Unicode; }
        }


        protected void PlayWScript(String code, String ext)
        {
            // Write the code to a .js/.vbs file.
            String     tempDir      = Path.GetTempPath();
            String     tempFileName = Guid.NewGuid().ToString() + ext;
            TextWriter textWriter   = new StreamWriter(tempDir + tempFileName, false, Encoding.Unicode);

            textWriter.Write(code);
            textWriter.Close();

            // Execute the .js file.
            Process scriptProc = new Process();
            scriptProc.StartInfo.FileName         = "wscript";
            scriptProc.StartInfo.WorkingDirectory = tempDir;
            scriptProc.StartInfo.Arguments        = tempFileName;
            scriptProc.StartInfo.UseShellExecute  = false;
            scriptProc.Start();
            scriptProc.WaitForExit(2000);
            scriptProc.Close();

            try
            {
                File.Delete(tempFileName);
            }
            catch (Exception)
            {
                Debug.Assert(false);
            }
        }


        protected void PlayFile(String code, String ext)
        {
            // Write the code to a .js/.vbs file.
            String     tempDir      = Path.GetTempPath();
            String     tempFileName = Guid.NewGuid().ToString() + ext;
            TextWriter textWriter   = new StreamWriter(tempDir + tempFileName, false, this.GeneratorEncoding);

            textWriter.Write(code);
            textWriter.Close();

            // Execute the .js file.
            Process scriptProc = new Process();
            scriptProc.StartInfo.FileName         = tempFileName;
            scriptProc.StartInfo.WorkingDirectory = tempDir;
            scriptProc.StartInfo.UseShellExecute  = true;
            scriptProc.Start();
            scriptProc.WaitForExit(2000);
            scriptProc.Close();

            try
            {
                File.Delete(tempFileName);
            }
            catch (Exception)
            {
                Debug.Assert(false);
            }
        }


        protected void PlayDotNet(String code, String language, String moduleName)
        {
            CodeDomProvider    codeProvider = CodeDomProvider.CreateProvider(language);
            CompilerParameters parameters   = new CompilerParameters();

            parameters.GenerateInMemory = true;
            parameters.ReferencedAssemblies.Add("Interop.SHDocVw.dll");
            parameters.ReferencedAssemblies.Add("Interop.OpenTwebstLib.dll");

            CompilerResults results = codeProvider.CompileAssemblyFromSource(parameters, code);
            if (results.Errors.Count > 0)
            {
                String error = "Line number: ";
                foreach (CompilerError compErr in results.Errors)
                {
                    error += compErr.Line + "\n";
                    error += "Error: " + compErr.ErrorNumber + " '" + compErr.ErrorText + "'";

                    throw new Exception(error);
                }
            }
            else
            {
                Module[]   mods = results.CompiledAssembly.GetModules(false);
                Type       type = mods[0].GetType(moduleName);
                MethodInfo main = type.GetMethod("Main");

                // Play .Net macro in new thread thus don't block the UI.
                Thread playThread = new Thread(this.PlayMain);
                playThread.Start(main);
            }
        }


        private void PlayMain(Object mainObj)
        {
            try
            {
                MethodInfo main = (MethodInfo)mainObj;
                main.Invoke(null, null);
            }
            catch (Exception)
            {
                // This can happen if the user choose not to continue evaluation.
            }
        }


        protected String IdentCode(String code, int spaces)
        {
            String result = "";

            using (StringReader reader = new StringReader(code))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    result += "".PadLeft(spaces) + line + "\r\n";
                }
            }

            return result;
        }


        protected virtual String EncodeAttrName(String attrName)
        {
            if (attrName == "innertext")
            {
                return "uiName";
            }
            else
            {
                return attrName;
            }
        }


        protected virtual String EncodeHtmlTagName(String tagName)
        {
            return tagName;
        }


        protected virtual int EncodeIndex(int index)
        {
            return index;
        }


        protected abstract String EscapeStr(String source);


        protected String BACK_NAVIGATION_STATEMENT;
        protected String FORWARD_NAVIGATION_STATEMENT;
        protected String START_UP_STATEMENT;
        protected String CLICK_NO_INDEX_STATEMENT;
        protected String CLICK_NO_INDEX_NO_ATTR_STATEMENT;
        protected String CLICK_STATEMENT;
        protected String CLICK_NO_ATTR_STATEMENT;
        protected String TEXT_CHANGED_NO_INDEX_STATEMENT;
        protected String TEXT_CHANGED_NO_INDEX_NO_ATTR_STATEMENT;
        protected String TEXT_CHANGED_STATEMENT;
        protected String TEXT_CHANGED_NO_ATTR_STATEMENT;
        protected String TEXT_CHANGED_ON_FILE_IE8_COMMENT;
        protected String SELECT_MULTIPLE_DECLARATION;
        protected String SELECT_MULTIPLE_FIRST_ITEM_STATEMENT;
        protected String SELECT_MULTIPLE_NO_INDEX_STATEMENT;
        protected String SELECT_MULTIPLE_NO_INDEX_NO_ATTR_STATEMENT;
        protected String SELECT_MULTIPLE_STATEMENT;
        protected String SELECT_MULTIPLE_NO_ATTR_STATEMENT;
        protected String ADD_SELECTION_STATEMENT;
        protected String SELECT_NO_INDEX_STATEMENT;
        protected String SELECT_NO_INDEX_NO_ATTR_STATEMENT;
        protected String SELECT_STATEMENT;
        protected String SELECT_NO_ATTR_STATEMENT;
        protected String CLICK_ACTION_CALL       = "Click()";
        protected String RIGHT_CLICK_ACTION_CALL = "RightClick()";
        protected String CHECK_ACTION_CALL       = "Check()";
        protected String UNCHECK_ACTION_CALL     = "Uncheck()";


        protected String IEVersion
        {
            get
            {
                if (BaseLanguageGenerator.ieVersion == null)
                {
                    BaseLanguageGenerator.ieVersion = CoreWrapper.Instance.IEVersion;
                }

                return BaseLanguageGenerator.ieVersion;
            }
        }

        private static String ieVersion = null;
        
    }
}
