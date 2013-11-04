namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    public class ScriptErrorInfo
    {
        public ScriptErrorInfo()
        {
        }

        internal ScriptErrorInfo(IActiveScriptError error, IDictionary<ulong, ScriptInfo> scripts)
        {
            ExtractExcepInfo(error);
            ulong cookie = ExtractSourcePosition<uint>(error.GetSourcePosition);
            ExtractSourceText(error, cookie, scripts);
        }

        internal ScriptErrorInfo(IActiveScriptError64 error, IDictionary<ulong, ScriptInfo> scripts)
        {
            ExtractExcepInfo(error);
            ulong cookie = ExtractSourcePosition<ulong>(error.GetSourcePosition64);
            ExtractSourceText(error, cookie, scripts);
        }

        public int LineNumber { get; set; }
        public int ColumnNumber { get; set; }
        public string LineText { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public int ErrorNumber { get; set; }
        public string HelpFile { get; set; }
        public int HelpContext { get; set; }
        public string ScriptName { get; set; }

        // TODO: Remove from this class.
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

        private void ExtractExcepInfo(IActiveScriptError error)
        {
            EXCEPINFO excep;
            excep = error.GetExceptionInfo();

            /*
             * http://msdn.microsoft.com/en-us/library/windows/desktop/ms221133(v=vs.85).aspx
             * 
             * wCode
             * The error code. Error codes should be greater than 1000. 
             * Either this field or the scode field must be filled in; the other must be set to 0.
             * 
             * scode
             * A return value that describes the error. Either this field or wCode (but not both) must be filled in; 
             * the other must be set to 0. (16-bit Windows versions only.)
             * 
             */
            ErrorNumber = excep.scode != 0 ? excep.scode : excep.wCode;
            Source = excep.bstrSource;
            HelpFile = excep.bstrHelpFile;
            HelpContext = excep.dwHelpContext;
            Description = excep.bstrDescription;
        }

        private void ExtractSourceText(IActiveScriptError error, ulong cookie, IDictionary<ulong, ScriptInfo> scripts)
        {
            if (error == null || scripts == null)
            {
                return;
            }

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

        private TCookie ExtractSourcePosition<TCookie>(GetSourcePositionDelegate<TCookie> getSourcePosition) 
            where TCookie : struct
        {
            TCookie cookie;
            uint lineNum;
            int colNum;

            getSourcePosition(out cookie, out lineNum, out colNum);

            LineNumber = (int)lineNum;
            ColumnNumber = colNum < 0 ? 0 : colNum;

            return cookie;
        }

        private delegate void GetSourcePositionDelegate<TCookie>(out TCookie cookie, out uint lineNum, out int columnNumber)
            where TCookie : struct;
    }
}
