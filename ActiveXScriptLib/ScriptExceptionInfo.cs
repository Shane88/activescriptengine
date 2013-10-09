namespace ActiveXScriptLib
{
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    public class ScriptExceptionInfo
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

        public ScriptExceptionInfo()
        {
        }

        public ScriptExceptionInfo(IActiveScriptError error)
        {
            EXCEPINFO excep;
            excep = error.GetExceptionInfo();

            Description = excep.bstrDescription;
            ErrorNumber = excep.scode;
            ErrorNumberShort = excep.wCode;
            Source = excep.bstrSource;
            HelpFile = excep.bstrHelpFile;
            HelpContext = excep.dwHelpContext;

            int lineNum = 0, cookie = 0, colNum = 0;
            error.GetSourcePosition(out cookie, out lineNum, out colNum);

            LineNumber = lineNum;
            ColumnNumber = colNum;

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
