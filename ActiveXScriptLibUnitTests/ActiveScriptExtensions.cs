namespace ActiveXScriptLib
{
    using System;

    public static class ActiveScriptExtensions
    {
        public static string DebugDump(this ScriptErrorInfo errorInfo)
        {
            return string.Format(
                "Error in {0} on Ln {1} Col {2}, {3}{4}Error Code:{5}, Text:{6}",
                errorInfo.ScriptName ?? "Unnamed Script",
                errorInfo.LineNumber,
                errorInfo.ColumnNumber,
                errorInfo.Description,
                Environment.NewLine,
                errorInfo.ErrorNumber,
                errorInfo.LineText);
        }
    }
}
