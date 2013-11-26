﻿namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    /// <summary>
    /// Implemented by the host to create a site for the Windows Script engine. 
    /// Usually, this site will be associated with the container of all the objects that are visible
    /// to the script (for example, the ActiveX Controls). Typically, this container will correspond
    /// to the document or page being viewed. Microsoft Internet Explorer, for example, would create
    /// such a container for each HTML page being displayed. Each ActiveX control (or other automation object)
    /// on the page, and the scripting engine itself, would be enumerable within this container.
    /// <see cref="http://msdn.microsoft.com/en-us/library/z70w3w6a(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("db01a1e3-a42b-11cf-8f20-00805f2cd064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptSite
    {
        /// <summary>
        /// Retrieves the locale identifier associated with the host's user interface. 
        /// The scripting engine uses the identifier to ensure that error strings and other 
        /// user-interface elements generated by the engine appear in the appropriate language.
        /// <see cref="http://msdn.microsoft.com/en-us/library/4xkx4eyw(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="lcid">Address of a variable that receives the locale identifier for 
        /// user-interface elements displayed by the scripting engine.</param>
        uint GetLCID();

        /// <summary>
        /// Allows the scripting engine to obtain information about an item added with the IActiveScript::AddNamedItem method.
        /// This method retrieves only the information indicated by the dwReturnMask parameter; 
        /// this improves performance. For example, if an ITypeInfo interface is not needed for an item, 
        /// it is simply not specified in dwReturnMask.
        /// <see cref="http://msdn.microsoft.com/en-us/library/442d1dx8(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="name">The name associated with the item, as specified in the IActiveScript::AddNamedItem method.</param>
        /// <param name="mask">A bit mask specifying what information about the item should be returned.
        /// The scripting engine should request the minimum amount of information possible because 
        /// some of the return parameters (for example, ITypeInfo) can take considerable time to load or generate.</param>
        /// <param name="pUnkItem">Address of a variable that receives a pointer to the IUnknown 
        /// interface associated with the given item. The scripting engine can use the 
        /// IUnknown::QueryInterface method to obtain the IDispatch interface for the item. 
        /// This parameter receives NULL if dwReturnMask does not include the SCRIPTINFO_IUNKNOWN value.
        /// Also, it receives NULL if there is no object associated with the item name; 
        /// this mechanism is used to create a simple class when the named item was added with the 
        /// SCRIPTITEM_CODEONLY flag set in the IActiveScript::AddNamedItem method.</param>
        /// <param name="pTypeInfo">Address of a variable that receives a pointer to the ITypeInfo 
        /// interface associated with the item. This parameter receives NULL if dwReturnMask does 
        /// not include the SCRIPTINFO_ITYPEINFO value, or if type information is not available for
        /// this item. If type information is not available, the object cannot source events, and 
        /// name binding must be realized with the IDispatch::GetIDsOfNames method. Note that the 
        /// ITypeInfo interface retrieved describes the item's coclass (TKIND_COCLASS) because the 
        /// object may support multiple interfaces and event interfaces. If the item supports the 
        /// IProvideMultipleTypeInfo interface, the ITypeInfo interface retrieved is the same as the
        /// index zero ITypeInfo that would be obtained using the IProvideMultipleTypeInfo::GetInfoOfIndex method.</param>
        void GetItemInfo(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string name,
            [In] ScriptInfoFlags mask,
            [In] [Out] ref IntPtr pUnkItem,
            [In] [Out] ref IntPtr pTypeInfo
        );

        /// <summary>
        /// Retrieves a host-defined string that uniquely identifies the current document version. 
        /// If the related document has changed outside the scope of Windows Script 
        /// (as in the case of an HTML page being edited with Notepad), the scripting engine can 
        /// save this along with its persisted state, forcing a recompile the next time the script is loaded.
        /// If E_NOTIMPL is returned, the scripting engine should assume that the script is in sync with the document.
        /// <see cref="http://msdn.microsoft.com/en-us/library/a80a8e1d(v=vs.94).aspx"/>
        /// </summary>
        /// <returns>The host-defined document version string.</returns>
        string GetDocVersionString();

        /// <summary>
        /// Informs the host that the script has completed execution.
        /// The scripting engine calls this method before the call to the IActiveScriptSite::OnStateChange 
        /// method, with the SCRIPTSTATE_INITIALIZED flag set, is completed. This method can be used 
        /// to return completion status and results to the host. Note that many script languages,
        /// which are based on sinking events from the host, have life spans that are defined by the host. 
        /// In this case, this method may never be called.
        /// <see cref="http://msdn.microsoft.com/en-us/library/03w8sd4a(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="pVarResult">Address of a variable that contains the script result, 
        /// or NULL if the script produced no result.</param>
        /// <param name="excepInfo">Address of an EXCEPINFO structure that contains exception 
        /// information generated when the script terminated, or NULL if no exception was generated.</param>
        void OnScriptTerminate(
            [In] IntPtr pVarResult,
            [In] ref EXCEPINFO excepInfo
        );

        /// <summary>
        /// Informs the host that the scripting engine has changed states.
        /// <see cref="http://msdn.microsoft.com/en-us/library/bd7cy77w(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="state">Value that indicates the new script state. 
        /// See the IActiveScript::GetScriptState method for a description of the states.</param>
        void OnStateChange(
            [In] ScriptState state
        );

        /// <summary>
        /// Informs the host that an execution error occurred while the engine was running the script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/shbz8x82(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="error">Address of the error object's IActiveScriptError interface. 
        /// A host can use this interface to obtain information about the execution error.</param>
        void OnScriptError(
            [In] IActiveScriptError error
        );

        /// <summary>
        /// Informs the host that the scripting engine has begun executing the script code.
        /// <see cref="http://msdn.microsoft.com/en-us/library/9c1cww48(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>The scripting engine must call this method on every entry or reentry into the 
        /// scripting engine. For example, if the script calls an object that then fires an event 
        /// handled by the scripting engine, the scripting engine must call IActiveScriptSite::OnEnterScript
        /// before executing the event, and must call the IActiveScriptSite::OnLeaveScript method 
        /// after executing the event but before returning to the object that fired the event. 
        /// Calls to this method can be nested. Every call to this method requires a corresponding 
        /// call to IActiveScriptSite::OnLeaveScript.</remarks>
        void OnEnterScript();

        /// <summary>
        /// Informs the host that the scripting engine has returned from executing script code.
        /// <see cref="http://msdn.microsoft.com/en-us/library/ebahs1kk(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>The scripting engine must call this method before returning control to a calling
        /// application that entered the scripting engine. For example, if the script calls an object
        /// that then fires an event handled by the scripting engine, the scripting engine must call
        /// the IActiveScriptSite::OnEnterScript method before executing the event, and must call 
        /// IActiveScriptSite::OnLeaveScript after executing the event before returning to the object
        /// that fired the event. Calls to this method can be nested. Every call to 
        /// IActiveScriptSite::OnEnterScript requires a corresponding call to this method.</remarks>
        void OnLeaveScript();
    }
}
