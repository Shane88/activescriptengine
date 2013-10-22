namespace ActiveXScriptLib
{
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    public class ScriptErrorInfo
    {
        public ulong LineNumber { get; set; }
        public ulong ColumnNumber { get; set; }
        public string LineText { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public int ErrorNumber { get; set; }
        public int ErrorNumberShort { get; set; }
        public string HelpFile { get; set; }
        public int HelpContext { get; set; }

        public ScriptErrorInfo()
        {
        }

        public ScriptErrorInfo(IActiveScriptError error)
        {
            EXCEPINFO excep;
            excep = error.GetExceptionInfo();

            Description = excep.bstrDescription;
            ErrorNumber = excep.scode;
            ErrorNumberShort = excep.wCode;
            Source = excep.bstrSource;
            HelpFile = excep.bstrHelpFile;
            HelpContext = excep.dwHelpContext;

            uint cookie = 0; 
            uint lineNum = 0;
            int colNum = 0;
            error.GetSourcePosition(out cookie, out lineNum, out colNum);

            LineNumber = lineNum;
            ColumnNumber = colNum < 0 ? 0 : (ulong)colNum;

            try
            {
                LineText = error.GetSourceLineText();
            }
            catch
            {
            }
        }

        public ScriptErrorInfo(IActiveScriptError64 error)
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

            LineNumber = lineNum;
            ColumnNumber = colNum < 0 ? 0 : (ulong)colNum;

            try
            {
                LineText = error.GetSourceLineText();
            }
            catch
            {
            }
        }
    }
}
