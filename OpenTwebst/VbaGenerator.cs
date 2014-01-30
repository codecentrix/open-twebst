using System;
using System.Collections.Generic;
using System.Text;



namespace CatStudio
{
    class VbaGenerator : BaseLanguageGenerator
    {
        public VbaGenerator()
        {
            this.BACK_NAVIGATION_STATEMENT                  = "Call browser.nativeBrowser.GoBack()";
            this.FORWARD_NAVIGATION_STATEMENT               = "Call browser.nativeBrowser.GoForward()";
            this.START_UP_STATEMENT                         = "Dim core As ICore\n" +
                                                              "Set core = New OpenTwebstLib.core\n\n" +
                                                              "Dim browser As IBrowser\n" +
                                                              "Set browser = core.StartBrowser(\"{0}\")\n";
            this.CLICK_NO_INDEX_STATEMENT                   = "Call browser.FindElement(\"{0}\", \"{1}={2}\").{3}";
            this.CLICK_NO_INDEX_NO_ATTR_STATEMENT           = "Call browser.FindElement(\"{0}\", \"\").{1}";
            this.CLICK_STATEMENT                            = "Call browser.FindElement(\"{0}\", \"{1}={2}, index={3}\").{4}";
            this.CLICK_NO_ATTR_STATEMENT                    = "Call browser.FindElement(\"{0}\", \"index={1}\").{2}";
            this.TEXT_CHANGED_NO_INDEX_STATEMENT            = "Call browser.FindElement(\"{0}\", \"{1}={2}\").InputText(\"{3}\")";
            this.TEXT_CHANGED_NO_INDEX_NO_ATTR_STATEMENT    = "Call browser.FindElement(\"{0}\", \"\").InputText(\"{1}\")";
            this.TEXT_CHANGED_STATEMENT                     = "Call browser.FindElement(\"{0}\", \"{1}={2}, index={3}\").InputText(\"{4}\")";
            this.TEXT_CHANGED_NO_ATTR_STATEMENT             = "Call browser.FindElement(\"{0}\", \"index={1}\").InputText(\"{2}\")";
            this.TEXT_CHANGED_ON_FILE_IE8_COMMENT           = "'Because of new HTML 5 security specifications, IE8 - IE9 does not reveal the real local path of the file you have selected. You have to manually change the code";
            this.SELECT_MULTIPLE_DECLARATION                = "Dim s As IElement\n";
            this.SELECT_MULTIPLE_FIRST_ITEM_STATEMENT       = "s.Select(\"{0}\")";
            this.SELECT_MULTIPLE_NO_INDEX_STATEMENT         = "{0}s = browser.FindElement(\"{1}\", \"{2}={3}\")\n";
            this.SELECT_MULTIPLE_NO_INDEX_NO_ATTR_STATEMENT = "{0}s = browser.FindElement(\"{1}\", \"\")\n";
            this.SELECT_MULTIPLE_STATEMENT                  = "{0}s = browser.FindElement(\"{1}\", \"{2}={3}, index={4}\")\n";
            this.SELECT_MULTIPLE_NO_ATTR_STATEMENT          = "{0}s = browser.FindElement(\"{1}\", \"index={2}\")\n";
            this.ADD_SELECTION_STATEMENT                    = "\ns.AddSelection(\"{0}\")";
            this.SELECT_NO_INDEX_STATEMENT                  = "Call browser.FindElement(\"{0}\", \"{1}={2}\").Select(\"{3}\")";
            this.SELECT_NO_INDEX_NO_ATTR_STATEMENT          = "Call browser.FindElement(\"{0}\", \"\").Select(\"{1}\")";
            this.SELECT_STATEMENT                           = "Call browser.FindElement(\"{0}\", \"{1}={2}, index={3}\").Select(\"{4}\")";
            this.SELECT_NO_ATTR_STATEMENT                   = "Call browser.FindElement(\"{0}\", \"index={1}\").Select(\"{2}\")";

            // In VBA paranthesis are not needed.
            this.CLICK_ACTION_CALL   = "Click";
            this.CHECK_ACTION_CALL   = "Check";
            this.UNCHECK_ACTION_CALL = "Uncheck";
        }


        protected override String EscapeStr(String source)
        {
            if (source == null)
            {
                return null;
            }

            String result = source.Replace("\"", "\"\"").Replace("\r\n", "\" & Microsoft.VisualBasic.ControlChars.CrLf & \"").Replace("\r", "\" & Microsoft.VisualBasic.ControlChars.Cr & \"").Replace("\n", "\" & Microsoft.VisualBasic.ControlChars.Lf & \"");
            return result;
        }


        internal override void Play(String code)
        {
            throw new Exception("Play is not supported for VBA language!\nYou need to manually include generated code into Excel/Word macro.");
        }


        internal override String FileExt
        {
            get { return ".bas"; }
        }


        internal override String DecorateCode(String code)
        {
            return String.Format(this.vbaDecoration, IdentCode(code, 4));
        }


        public override string ToString()
        {
            return "VBA";
        }


        private readonly String vbaDecoration = 
@"'Add Open Twebst Type Library in Tools/References menu of the VBA editor.
Sub OpenTwebstMacro
{0}
End Sub
";
    }
}
