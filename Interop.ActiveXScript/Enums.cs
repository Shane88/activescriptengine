namespace Interop.ActiveXScript
{
    using System;

    [Flags]
    public enum ScriptInfoFlags : uint
    {
        None = 0,
        IUnknown = 0x00000001,
        ITypeInfo = 0x00000002
    }

    [Flags]
    public enum ScriptInterruptFlags : uint
    {
        None = 0,
        Debug = 0x00000001,
        RaiseException = 0x00000002
    }

    /// <summary>
    /// Flags associated with a named item.
    /// http://msdn.microsoft.com/en-us/library/s8eyc3sh(v=vs.94).aspx
    /// </summary>
    [Flags]
    public enum ScriptItemFlags : uint
    {
        None = 0,

        /// <summary>
        /// Indicates that the item's name is available in the name space of the script, 
        /// allowing access to the properties, methods, and events of the item. 
        /// By convention the properties of the item include the item's children; therefore,
        /// all child object properties and methods (and their children, recursively) will be accessible.
        /// </summary>
        IsVisible = 0x00000002,

        /// <summary>
        /// Indicates that the item sources events that the script can sink. Child objects (properties of 
        /// the object that are in themselves objects) can also source events to the script. This is not
        /// recursive, but it provides a convenient mechanism for the common case, for example, of creating
        /// a container and all of its member controls.
        /// </summary>
        IsSource = 0x00000004,

        /// <summary>
        /// Indicates that the item is a collection of global properties and methods associated with the script. 
        /// Normally, a scripting engine would ignore the object name (other than for the purpose of using 
        /// it as a cookie for the IActiveScriptSite::GetItemInfo method, or for resolving explicit scoping)
        /// and expose its members as global variables and methods. This allows the host to extend the library
        /// (run-time functions and so on) available to the script. It is left to the scripting engine to deal
        /// with name conflicts (for example, when two SCRIPTITEM_GLOBALMEMBERS items have methods of the same name),
        /// although an error should not be returned because of this situation.
        /// </summary>
        GlobalMembers = 0x00000008,

        /// <summary>
        /// Indicates that the item should be saved if the scripting engine is saved. Similarly, setting this 
        /// flag indicates that a transition back to the initialized state should retain the item's name and
        /// type information (the scripting engine must, however, release all pointers to interfaces on the actual object).
        /// </summary>
        IsPersistent = 0x00000040,

        /// <summary>
        /// Indicates that the named item represents a code-only object, and that the host has no 
        /// IUnknown to be associated with this code-only object. The host only has a name for 
        /// this object. In object-oriented languages such as C++, this flag would create a class.
        /// Not all languages support this flag.
        /// </summary>
        CodeOnly = 0x00000200,

        /// <summary>
        /// Indicates that the item is simply a name being added to the script's name space, 
        /// and should not be treated as an item for which code should be associated. 
        /// For example, without this flag being set, VBScript will create a separate 
        /// module for the named item, and C++ might create a separate wrapper class for the named item.
        /// </summary>
        NoCode = 0x00000400
    }

    /// <summary>
    /// Specifies the state of a scripting engine. This enumeration is used by the 
    /// IActiveScript::GetScriptState, IActiveScript::SetScriptState, and IActiveScriptSite::OnStateChange methods.
    /// <see cref="http://msdn.microsoft.com/en-us/library/f7z7cxxa(v=vs.94).aspx"/>
    /// </summary>
    public enum ScriptState : uint
    {
        /// <summary>
        /// Script has just been created, but has not yet been initialized using an IPersist* 
        /// interface and IActiveScript::SetScriptSite.
        /// </summary>
        Uninitialized = 0,

        /// <summary>
        /// Script has been initialized, but is not running (connecting to other objects or sinking events)
        /// or executing any code. Code can be queried for execution by calling the IActiveScriptParse::ParseScriptText method.
        /// </summary>
        Initialized = 5,

        /// <summary>
        /// Script can execute code, but is not yet sinking the events of objects added by the 
        /// IActiveScript::AddNamedItem method.
        /// </summary>
        Started = 1,

        /// <summary>
        /// Script is loaded and connected for sinking events.
        /// </summary>
        Connected = 2,

        /// <summary>
        /// Script is loaded and has a run-time execution state, but is temporarily disconnected from sinking events.
        /// </summary>
        Disconnected = 3,

        /// <summary>
        /// Script has been closed. The scripting engine no longer works and returns errors for most methods.
        /// </summary>
        Closed = 4
    }

    public enum ScriptThreadState : uint
    {
        NotInScript = 0,
        Running = 1
    }

    [Flags]
    public enum ScriptTypeLibFlags : uint
    {
        None = 0,
        IsControl = 0x00000010,
        IsPersistent = 0x00000040
    }

    [Flags]
    public enum ScriptTextFlags : uint
    {
        None = 0,
        DelayExecution = 0x00000001,

        /// <summary>
        /// Indicates that the script text should be visible (and, therefore, callable by name) as 
        /// a global method in the name space of the script.
        /// </summary>
        IsVisible = 0x00000002,

        /// <summary>
        /// If the distinction between a computational expression and a statement is important but 
        /// syntactically ambiguous in the script language, this flag specifies that the scriptlet 
        /// is to be interpreted as an expression, rather than as a statement or list of statements.
        /// By default, statements are assumed unless the correct choice can be determined from the
        /// syntax of the scriptlet text.
        /// </summary>
        IsExpression = 0x00000020,

        /// <summary>
        /// Indicates that the code added during this call should be saved if the scripting engine 
        /// is saved (for example, through a call to IPersist*::Save), or if the scripting engine 
        /// is reset by way of a transition back to the initialized state.
        /// </summary>
        IsPersistent = 0x00000040,
        HostManagesSource = 0x00000080,
        IsXDomain = 0x00000100
    }

    public enum ScriptGCType
    {
        Normal = 0,
        Exhaustive = 1
    }
}
