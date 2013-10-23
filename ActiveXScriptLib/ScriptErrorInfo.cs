namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    public class ScriptErrorInfo
    {
        public int LineNumber { get; set; }
        public int ColumnNumber { get; set; }
        public string LineText { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public int ErrorNumber { get; set; }
        public int ErrorNumberShort { get; set; }
        public string HelpFile { get; set; }
        public int HelpContext { get; set; }
        public string ScriptName { get; set; }

        public ScriptErrorInfo()
        {
        }

        internal ScriptErrorInfo(IDictionary<uint, ScriptInfo> scripts, IActiveScriptError error)
        {
            EXCEPINFO excep;
            excep = error.GetExceptionInfo();

            Description      = excep.bstrDescription;
            ErrorNumber      = excep.scode;
            ErrorNumberShort = excep.wCode;
            Source           = excep.bstrSource;
            HelpFile         = excep.bstrHelpFile;
            HelpContext      = excep.dwHelpContext;

            uint cookie  = 0;
            uint lineNum = 0;
            int colNum   = 0;
            error.GetSourcePosition(out cookie, out lineNum, out colNum);

            LineNumber = (int)lineNum;
            ColumnNumber = colNum < 0 ? 0 : colNum;

            ScriptInfo scriptInfo = null;

            if (scripts.ContainsKey(cookie))
            {
                scriptInfo = scripts[cookie];
                ScriptName = scriptInfo.ScriptName;
            }

            try
            {
                LineText = error.GetSourceLineText();
            }
            catch
            {
                if (scriptInfo != null) 
                {
                    string[] lines = scriptInfo.Code.Split(new string[] { Environment.NewLine },
                        StringSplitOptions.None);

                    if (lines != null && lines.Length > 0 && LineNumber - 1 <= lines.Length)
                    {
                        LineText = lines[LineNumber - 1];
                    }
                }
            }
        }

        internal ScriptErrorInfo(IActiveScriptError64 error)
        {
            EXCEPINFO excep;
            excep = error.GetExceptionInfo();

            Description = excep.bstrDescription;
            ErrorNumber = excep.scode;
            ErrorNumberShort = excep.wCode;
            Source = excep.bstrSource;
            HelpFile = excep.bstrHelpFile;
            HelpContext = excep.dwHelpContext;

            ulong cookie = 0;
            uint lineNum = 0;
            int colNum = 0;
            error.GetSourcePosition64(out cookie, out lineNum, out colNum);

            LineNumber = (int)lineNum;
            ColumnNumber = colNum < 0 ? 0 : colNum;

            try
            {
                LineText = error.GetSourceLineText();
            }
            catch
            {
            }
        }

        public string DebugDump()
        {
            return string.Format(
                "Error in {0} on Ln {1} Col {2}, {3}{4}Error Code:{5}, Text:{6}",
                ScriptName ?? "Unnamed Script", LineNumber, ColumnNumber, Description, 
                Environment.NewLine, ErrorNumber, LineText);
        }

        public override string ToString()
        {
            return this.Description;
        }
    }
}
