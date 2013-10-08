namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    /// <summary>
    /// If the Windows Script engine allows raw text code scriptlets to be added to the script or 
    /// allows expression text to be evaluated at run time, it implements the IActiveScriptParse interface.
    /// For interpreted scripting languages that have no independent authoring environment, 
    /// such as VBScript, this provides an alternate mechanism (other than IPersist*) to get script
    /// code into the scripting engine, and to attach script fragments to various object events.
    /// <see cref="http://msdn.microsoft.com/en-us/library/f2822wbt(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("bb1a2ae2-a4f9-11cf-8f20-00805f2cd064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptParse32
    {
        /// <summary>
        /// Initializes the scripting engine.
        /// <see cref="http://msdn.microsoft.com/en-us/library/5z520akb(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>Before the scripting engine can be used, one of the following methods must be 
        /// called: IPersist*::Load, IPersist*::InitNew, or IActiveScriptParse::InitNew. 
        /// The semantics of this method are identical to IPersistStreamInit::InitNew, in that this
        /// method tells the scripting engine to initialize itself. Note that it is not valid to call
        /// both IPersist*::InitNew or IActiveScriptParse::InitNew and IPersist*::Load, nor is it 
        /// valid to call IPersist*::InitNew, IActiveScriptParse::InitNew, or IPersist*::Load more
        /// than once.</remarks>
        void InitNew();

        /// <summary>
        /// Adds a code scriptlet to the script. This method is used in environments where the persistent
        /// state of the script is intertwined with the host document and the host is responsible 
        /// for restoring the script, rather than through an IPersist* interface. 
        /// The primary examples are HTML scripting languages that allow scriptlets of code embedded
        /// in the HTML document to be attached to intrinsic events 
        /// (for instance, ONCLICK="button1.text='Exit'").
        /// <see cref="http://msdn.microsoft.com/en-us/library/80dz62xx(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="defaultName">Address of a default name to associate with the scriptlet. 
        /// If the scriptlet does not contain naming information (as in the ONCLICK example above),
        /// this name will be used to identify the scriptlet. 
        /// If this parameter is NULL, the scripting engine manufactures a unique name, if necessary.</param>
        /// <param name="code">Address of the scriptlet text to add. The interpretation of this 
        /// string depends on the scripting language.</param>
        /// <param name="itemName">Address of a buffer that contains the item name associated with 
        /// this scriptlet. This parameter, in addition to pstrSubItemName, identifies the object
        /// for which the scriptlet is an event handler.</param>
        /// <param name="subItemName">Address of a buffer that contains the name of a subobject of 
        /// the named item with which this scriptlet is associated; this name must be found in the 
        /// named item's type information. This parameter is NULL if the scriptlet is to be associated
        /// with the named item instead of a subitem. This parameter, in addition to pstrItemName, 
        /// identifies the specific object for which the scriptlet is an event handler.</param>
        /// <param name="eventName">Address of a buffer that contains the name of the event for 
        /// which the scriptlet is an event handler.</param>
        /// <param name="delimiter">Address of the end-of-scriptlet delimiter. When the pstrCode parameter
        /// is parsed from a stream of text, the host typically uses a delimiter, such as two single
        /// 4quotation marks (''), to detect the end of the scriptlet. This parameter specifies the
        /// delimiter that the host used, allowing the scripting engine to provide some conditional
        /// primitive preprocessing (for example, replacing a single quotation mark ['] with two
        /// single quotation marks for use as a delimiter). Exactly how (and if) the scripting engine
        /// makes use of this information depends on the scripting engine. Set this parameter to NULL
        /// if the host did not use a delimiter to mark the end of the scriptlet.</param>
        /// <param name="sourceContext">Application-defined value that is used for debugging purposes.</param>
        /// <param name="startingLineNumber">Zero-based value that specifies which line the parsing will begin at.</param>
        /// <param name="flags">Flags associated with the scriptlet.</param>
        /// <param name="name">Actual name used to identify the scriptlet. This is to be in order 
        /// of preference: a name explicitly specified in the scriptlet text, the default name provided
        /// in pstrDefaultName, or a unique name synthesized by the scripting engine.</param>
        /// <param name="excepInfo">Address of a structure containing exception information. 
        /// This structure should be filled in if DISP_E_EXCEPTION is returned.</param>
        void AddScriptlet(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string defaultName,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string code,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string itemName,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string subItemName,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string eventName,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string delimiter,
            [In] uint sourceContext,
            [In] uint startingLineNumber,
            [In] ScriptTextFlags flags,
            [Out] [MarshalAs(UnmanagedType.BStr)] out string name,
            [Out] out EXCEPINFO excepInfo
        );

        /// <summary>
        /// Parses the given code scriptlet, adding declarations into the namespace and evaluating 
        /// code as appropriate.
        /// <see cref="http://msdn.microsoft.com/en-us/library/tch4w30x(v=vs.94).aspx"/>
        /// </summary>
        /// <remarks>If the scripting engine is in the initialized state, no code will actually be 
        /// evaluated during this call; rather, such code is queued and executed when the scripting
        /// engine is transitioned into (or through) the started state. Because execution is not
        /// allowed in the initialized state, it is an error to call this method with the 
        /// SCRIPTTEXT_ISEXPRESSION flag when in the initialized state. The scriptlet can be an 
        /// expression, a list of statements, or anything allowed by the script language. 
        /// For example, this method is used in the evaluation of the HTML &lt;SCRIPT&gt; tag, which allows
        /// statements to be executed as the HTML page is being constructed, rather than just compiling
        /// them into the script state. The code passed to this method must be a valid, complete portion
        /// of code. For example, in VBScript it is illegal to call this method once with Sub Function(x)
        /// and then a second time with End Sub. The parser must not wait for the second call to complete
        /// the subroutine, but rather must generate a parse error because a subroutine declaration
        /// was started but not completed. 
        /// For more information about script states, see the Script Engine States section of 
        /// Windows Script Engines.</remarks>
        /// <param name="code">Address of the scriptlet text to evaluate. 
        /// The interpretation of this string depends on the scripting language.</param>
        /// <param name="itemName">Address of the item name that gives the context in which the 
        /// scriptlet is to be evaluated. If this parameter is NULL, 
        /// the code is evaluated in the scripting engine's global context.</param>
        /// <param name="context">Address of the context object. This object is reserved for use in
        /// a debugging environment, where such a context may be provided by the debugger to represent
        /// an active run-time context. If this parameter is NULL, the engine uses pstrItemName to
        /// identify the context.</param>
        /// <param name="delimiter">Address of the end-of-scriptlet delimiter. When pstrCode is parsed
        /// from a stream of text, the host typically uses a delimiter, such as two single quotation
        /// marks (''), to detect the end of the scriptlet. This parameter specifies the delimiter 
        /// that the host used, allowing the scripting engine to provide some conditional primitive
        /// preprocessing (for example, replacing a single quotation mark ['] with two single 
        /// quotation marks for use as a delimiter). Exactly how (and if) the scripting engine makes
        /// use of this information depends on the scripting engine. Set this parameter to NULL if
        /// the host did not use a delimiter to mark the end of the scriptlet.</param>
        /// <param name="sourceContext">Cookie used for debugging purposes.</param>
        /// <param name="startingLineNumber">Zero-based value that specifies which line the parsing 
        /// will begin at.</param>
        /// <param name="flags">Flags associated with the scriptlet.</param>
        /// <param name="pVarResult">Address of a buffer that receives the results of scriptlet 
        /// processing, or NULL if the caller expects no result 
        /// (that is, the SCRIPTTEXT_ISEXPRESSION value is not set).</param>
        /// <param name="excepInfo">Address of a structure that receives exception information.
        /// This structure is filled if IActiveScriptParse::ParseScriptText returns DISP_E_EXCEPTION.</param>
        void ParseScriptText(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string code,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string itemName,
            [In] [MarshalAs(UnmanagedType.IUnknown)] object context,
            [In] [MarshalAs(UnmanagedType.LPWStr)] string delimiter,
            [In] uint sourceContext,
            [In] uint startingLineNumber,
            [In] ScriptTextFlags flags,
            [Out] object pVarResult,
            [Out] out EXCEPINFO excepInfo
        );
    }
}
