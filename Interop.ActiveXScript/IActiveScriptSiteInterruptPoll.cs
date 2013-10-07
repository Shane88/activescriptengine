namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// The IActiveScriptSiteInterruptPoll interface allows a host to specify that a script should terminate.
    /// <see cref="http://msdn.microsoft.com/en-us/library/sc7h6s9s(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("539698a0-cdca-11cf-a5eb-00aa0047a063")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptSiteInterruptPoll
    {
        /// <summary>
        /// Allows a host to specify that a script should terminate.
        /// <see cref="http://msdn.microsoft.com/en-us/library/839hsy0d(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>The hosted script should terminate unless the return value of the QueryContinue 
        /// method is S_OK. A return value of S_FALSE indicates that the host explicitly requests 
        /// that the script terminate. A multithreaded host may use the 
        /// IActiveScript::InterruptScriptThread method to terminate a script.</remarks>
        /// <returns>The method returns an HRESULT. Possible values include, 
        /// but are not limited to, those in the following table.
        /// S_OK 
        /// The call succeeded and the host permits the script to continue running.
        /// S_FALSE
        /// The call succeeded and the host requests that the script terminate.</returns>
        [PreserveSig]
        uint QueryContinue();
    }
}
