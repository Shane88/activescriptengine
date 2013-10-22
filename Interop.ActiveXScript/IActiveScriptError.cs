namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    /// <summary>
    /// An object implementing this interface is passed to the IActiveScriptSite::OnScriptError 
    /// method whenever the scripting engine encounters an unhandled error. The host then calls
    /// methods on this object to obtain information about the error that occurred.
    /// <see cref="http://msdn.microsoft.com/en-us/library/b17d71fw(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("eae1ba61-a4ed-11cf-8f20-00805f2cd064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptError
    {

        /// <summary>
        /// Retrieves information about an error that occurred while the scripting engine was running a script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/0fdy4f25(v=vs.94).aspx"/>
        /// </summary>
        /// <returns>An EXCEPINFO structure that receives error information.</returns>
        EXCEPINFO GetExceptionInfo();

        /// <summary>
        /// Retrieves the location in the source code where an error occurred while the scripting 
        /// engine was running a script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/xdf7c475(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="sourceContext">Address of a variable that receives a cookie that identifies
        /// the context. The interpretation of this parameter depends on the host application.</param>
        /// <param name="lineNumber">Address of a variable that receives the line number in the 
        /// source file where the error occurred.</param>
        /// <param name="position">Address of a variable that receives the character position 
        /// in the line where the error occurred.</param>
        void GetSourcePosition(
            [Out] out uint sourceContext,
            [Out] out uint lineNumber,
            [Out] out int position
        );

        /// <summary>
        /// Retrieves the line in the source file where an error occurred while a scripting engine 
        /// was running a script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/dwy967hz(v=vs.94).aspx"/>
        /// </summary>
        /// <returns>The line of source code in which the error occurred.</returns>
        string GetSourceLineText();
    }
}
