namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    /// <summary>
    /// Provides the methods necessary to initialize the scripting engine. 
    /// The scripting engine must implement the IActiveScript interface.
    /// <see cref="http://msdn.microsoft.com/en-us/library/ky29ffxd(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("bb1a2ae1-a4f9-11cf-8f20-00805f2cd064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScript
    {
        /// <summary>
        /// Informs the scripting engine of the IActiveScriptSite interface site provided by the host. 
        /// Call this method before any other IActiveScript interface methods is used.
        /// <see cref="http://msdn.microsoft.com/en-us/library/a3a8btht(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="site">Address of the host-supplied script site to be associated with this 
        /// instance of the scripting engine. The site must be uniquely assigned to this scripting engine instance;
        /// it cannot be shared with other scripting engines.</param>
        void SetScriptSite(
            [In] IActiveScriptSite site
        );

        /// <summary>
        /// Retrieves the site object associated with the Windows Script engine.
        /// <see cref="http://msdn.microsoft.com/en-us/library/876ah4t1(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="iid">Identifier of the requested interface.</param>        
        /// <returns>The interface pointer to the host's site object.</returns>
        [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 0)]
        object GetScriptSite(
            [In] ref Guid iid
        );

        /// <summary>
        /// Puts the scripting engine into the given state. This method can be called from non-base
        /// threads without resulting in a non-base callout to host objects or to the 
        /// IActiveScriptSite interface.
        /// <see cref="http://msdn.microsoft.com/en-us/library/xd9bt7sb(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="state">Sets the scripting engine to the given state. 
        /// Can be one of the values defined in the SCRIPTSTATE Enumeration enumeration.</param>
        void SetScriptState(
            [In] ScriptState state
        );

        /// <summary>
        /// Retrieves the current state of the scripting engine. This method can be called from 
        /// non-base threads without resulting in a non-base callout to host objects or to the 
        /// IActiveScriptSite interface.
        /// </summary>
        /// <param name="state">Address of a variable that receives a value defined in the 
        /// SCRIPTSTATE Enumeration enumeration. The value indicates the current state of the 
        /// scripting engine associated with the calling thread.</param>
        ScriptState GetScriptState();

        /// <summary>
        /// Causes the scripting engine to abandon any currently loaded script, lose its state, 
        /// and release any interface pointers it has to other objects, thus entering a closed state.
        /// Event sinks, immediately executed script text, and macro invocations that are already 
        /// in progress are completed before the state changes (use IActiveScript::InterruptScriptThread
        /// to cancel a running script thread). This method must be called by the creating host before
        /// the interface is released to prevent circular reference problems.
        /// <see cref="http://msdn.microsoft.com/en-us/library/6fds63c9(v=vs.94).aspx"/>
        /// </summary>
        void Close();

        /// <summary>
        /// Adds the name of a root-level item to the scripting engine's name space. 
        /// A root-level item is an object with properties and methods, an event source, or all three.
        /// <see cref="http://msdn.microsoft.com/en-us/library/s8eyc3sh(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="name">Address of a buffer that contains the name of the item as viewed 
        /// from the script. The name must be unique and persistable.</param>
        /// <param name="flags">Flags associated with an item. Can be a combination of these values</param>
        void AddNamedItem(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string name,
            [In] ScriptItemFlags flags
        );

        /// <summary>
        /// Adds a type library to the name space for the script. This is similar to the #include 
        /// directive in C/C++. It allows a set of predefined items such as class definitions, 
        /// typedefs, and named constants to be added to the run-time environment available to the script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/4hb95zwx(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="libid">CLSID of the type library to add.</param>
        /// <param name="major">Major version number.</param>
        /// <param name="minor">Minor version number.</param>
        /// <param name="flags">Option flags.</param>
        void AddTypeLib(
            [In] ref Guid libid,
            [In] uint major,
            [In] uint minor,
            [In] ScriptTypeLibFlags flags
        );

        /// <summary>
        /// Retrieves the IDispatch interface for the methods and properties associated with the
        /// currently running script.
        /// <see cref="http://msdn.microsoft.com/en-us/library/c5b043ew(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>Because methods and properties can be added by calling the IActiveScriptParse interface,
        /// the IDispatch interface returned by this method can dynamically support new methods and properties.
        /// Similarly, the IDispatch::GetTypeInfo method should return a new, unique ITypeInfo
        /// interface when methods and properties are added. Note, however, that language engines must
        /// not change the IDispatch interface in a way that is incompatible with any previous 
        /// ITypeInfo interface returned. That implies, for example, that DISPIDs will never be reused.</remarks>
        /// <param name="itemName">Address of a buffer that contains the name of the item for which
        /// the caller needs the associated dispatch object. If this parameter is NULL, the dispatch
        /// object contains as its members all of the global methods and properties defined by the script.
        /// Through the IDispatch interface and the associated ITypeInfo interface, the host can 
        /// invoke script methods or view and modify script variables.</param>
        /// <returns>The object associated with the script's global methods and properties. 
        /// If the scripting engine does not support such an object, NULL is returned.</returns>
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object GetScriptDispatch(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string itemName
        );

        /// <summary>
        /// Retrieves a scripting-engine-defined identifier for the currently executing thread. 
        /// The identifier can be used in subsequent calls to script thread execution-control methods
        /// such as the IActiveScript::InterruptScriptThread method.
        /// This method can be called from non-base threads without resulting in a non-base callout
        /// to host objects or to the IActiveScriptSite interface.
        /// <see cref="http://msdn.microsoft.com/en-us/library/7weths74(v=vs.94).aspx"/>
        /// </summary>
        /// <returns>The script thread identifier associated with the current thread. 
        /// The interpretation of this identifier is left to the scripting engine, but it can be 
        /// just a copy of the Windows thread identifier. If the Win32 thread terminates, 
        /// this identifier becomes unassigned and can subsequently be assigned to another thread.</returns>
        uint GetCurrentScriptThreadID();

        /// <summary>
        /// Retrieves a scripting-engine-defined identifier for the thread associated with the given Win32 thread.
        /// The retrieved identifier can be used in subsequent calls to script thread execution 
        /// control methods such as the IActiveScript::InterruptScriptThread method. This method
        /// can be called from non-base threads without resulting in a non-base callout to host 
        /// objects or to the IActiveScript::InterruptScriptThread interface.
        /// <see cref="http://msdn.microsoft.com/en-us/library/z5643a60(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="win32ThreadID">Thread identifier of a running Win32 thread in the current process.
        /// Use the IActiveScript::GetCurrentScriptThreadID function to retrieve the thread identifier
        /// of the currently executing thread.</param>
        /// <returns>The script thread identifier associated with the given Win32 thread. 
        /// The interpretation of this identifier is left to the scripting engine, 
        /// but it can be just a copy of the Windows thread identifier. 
        /// Note that if the Win32 thread terminates, this identifier becomes unassigned and may 
        /// subsequently be assigned to another thread.</returns>
        uint GetScriptThreadID(
            [In] uint win32ThreadID
        );

        /// <summary>
        /// Retrieves the current state of a script thread.
        /// <see cref="http://msdn.microsoft.com/en-us/library/6a6kx7bd(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>This method can be called from non-base threads without resulting in a non-base
        /// callout to host objects or to the IActiveScriptSite interface.</remarks>
        /// <param name="scriptThreadID">Identifier of the thread for which the state is desired, 
        /// or one of the following special thread identifiers:
        /// SCRIPTTHREADID_BASE
        /// The base thread; that is, the thread in which the scripting engine was instantiated.
        /// SCRIPTTHREADID_CURRENT
        /// The currently executing thread.</param>
        /// <returns>The state of the indicated thread.
        /// The state is indicated by one of the named constant values defined by the SCRIPTTHREADSTATE
        /// Enumeration enumeration. If this parameter does not identify the current thread, 
        /// the state may change at any time.</returns>
        ScriptThreadState GetScriptThreadState(
            [In] uint scriptThreadID
        );

        /// <summary>
        /// Interrupts the execution of a running script thread (an event sink, an immediate execution,
        /// or a macro invocation). This method can be used to terminate a script that is stuck 
        /// (for example, in an infinite loop). It can be called from non-base threads without 
        /// resulting in a non-base callout to host objects or to the IActiveScriptSite method.
        /// <see cref="http://msdn.microsoft.com/en-us/library/ecadx4td(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="scriptThreadID">Identifier of the thread to interrupt, or one of the following
        /// special thread identifier values:
        /// SCRIPTTHREADID_ALL
        /// All threads. The interrupt is applied to all script methods currently in progress. 
        /// Note that unless the caller has requested that the script be disconnected, the next scripted event causes script code to run again by calling the IActiveScript::SetScriptState method with the SCRIPTSTATE_DISCONNECTED or SCRIPTSTATE_INITIALIZED flag set.
        /// SCRIPTTHREADID_BASE
        /// The base thread; that is, the thread in which the scripting engine was instantiated.
        /// SCRIPTTHREADID_CURRENT
        /// The currently executing thread.</param>
        /// <param name="excepInfo">Address of an EXCEPINFO structure containing the error information
        /// that should be reported to the aborted script.</param>
        /// <param name="flags">Option flags associated with the interruption.</param>
        void InterruptScriptThread(
            [In] uint scriptThreadID,
            [In] ref EXCEPINFO excepInfo,
            [In] ScriptInterruptFlags flags
        );

        /// <summary>
        /// Clones the current scripting engine (minus any current execution state), returning a loaded
        /// scripting engine that has no site in the current thread. The properties of this new scripting
        /// engine will be identical to the properties the original scripting engine would be in if
        /// it were transitioned back to the initialized state.
        /// <see cref="http://msdn.microsoft.com/en-us/library/t4f5c129(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>The IActiveScript::Clone method is an optimization of IPersist*::Save, 
        /// CoCreateInstance, and IPersist*::Load, so the state of the new scripting engine should 
        /// be the same as if the state of the original scripting engine were saved and loaded into
        /// a new scripting engine. Named items are duplicated in the cloned scripting engine, 
        /// but specific object pointers for each item are forgotten and are obtained with the 
        /// IActiveScriptSite::GetItemInfo method. This allows an identical object model with per-thread
        /// entry points (an apartment model) to be used. This method is used for multithreaded server
        /// hosts that can run multiple instances of the same script. The scripting engine may return
        /// E_NOTIMPL, in which case the host can achieve the same result by duplicating the persistent
        /// state and creating a new instance of the scripting engine with an IPersist* interface.
        /// This method can be called from non-base threads without resulting in a non-base callout
        /// to host objects or to the IActiveScriptSite interface.</remarks>
        /// <returns>The IActiveScript interface of the cloned scripting engine. 
        /// The host must create a site and call the IActiveScript::SetScriptSite method on the
        /// new scripting engine before it will be in the initialized state and, therefore, usable.</returns>
        [return: MarshalAs(UnmanagedType.Interface)]
        IActiveScript Clone();
    }
}
