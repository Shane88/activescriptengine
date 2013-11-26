namespace ActiveXScriptLib
{
    using Interop.ActiveXScript;
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;
    
    [ComVisible(true)]
    public class ActiveScriptEngine : IDisposable
    {
        public event ScriptErrorOccurredDelegate ScriptErrorOccurred;

        public ScriptErrorInfo LastError { get; internal set; }

        internal IActiveScript activeScript;
        internal Dictionary<string, object> hostObjects;

        private ActiveScriptParse parser;
        private ActiveScriptSite scriptSite;

        private Dictionary<ulong, ScriptInfo> scripts;
        private string scriptToParse;

        public ActiveScriptEngine(string progID)
        {
            if (progID == null)
            {
                throw new ArgumentNullException("progID");
            }

            if (string.IsNullOrWhiteSpace(progID))
            {
                throw new ArgumentException("progID must not be empty or whitespace");
            }

            activeScript = CreateInstanceFromProgID<IActiveScript>(progID);

            if (activeScript == null)
            {
                throw new Exception("Unable to create an IActiveScript from " + progID);
            }

            parser = ActiveScriptParse.MakeActiveScriptParse(activeScript);

            parser.InitNew();

            scriptSite = new ActiveScriptSite(this);
            activeScript.SetScriptSite(scriptSite);

            hostObjects = new Dictionary<string, object>();
            scripts = new Dictionary<ulong, ScriptInfo>();
        }

        /// <summary>
        /// Gets a value indicating wether the script engine is running.
        /// </summary>
        public bool IsRunning
        {
            get
            {
                return activeScript.GetScriptState() == ScriptState.Started;
            }
        }

        /// <summary>
        /// Adds the specified object into the script context with the specified name.
        /// </summary>
        /// <param name="name">The name the object will be referenced as.</param>
        /// <param name="obj">The object to be added.</param>
        /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
        public void AddObject(string name, object obj)
        {
            if (hostObjects.ContainsKey(name))
            {
                throw new ArgumentException("object with that name already exists", "name");
            }
            
            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible);
        }

        /// <summary>
        /// Adds the specified object into the script context with the specified name.
        /// The methods on the object can be called without the name as if they were
        /// native functions availible in the script.
        /// </summary>
        /// <param name="name">The name the object can optionally be referenced as.</param>
        /// <param name="obj">The object to be added.</param>
        /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
        public void AddGlobalMemberObject(string name, object obj)
        {
            if (hostObjects.ContainsKey(name))
            {
                throw new ArgumentException("object with that name already exists", "name");
            }

            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible | ScriptItemFlags.GlobalMembers);
        }

        /// <summary>
        /// Returns if the script contains a host object of the specified name.
        /// </summary>
        /// <param name="name">The name of the object.</param>
        /// <returns>True if the script contains a host object of the specified name, otherwise false.</returns>
        public bool ScriptHasHostObject(string name)
        {
            return hostObjects.ContainsKey(name);
        }

        /// <summary>
        /// Gets the script host object of the specified name.
        /// </summary>
        /// <param name="name">The name of the object.</param>
        /// <returns>The script host object.</returns>
        /// <exception cref="ArgumentException">If the specified host object did not exist.</exception>
        public object GetScriptObject(string name)
        {
            if (!ScriptHasHostObject(name))
            {
                throw new ArgumentException("object with that name does not exist", "name");
            }

            return hostObjects[name];
        }

        /// <summary>
        /// Adds the specified code to the scripting engine.
        /// </summary>
        /// <param name="code">The code to be added.</param>
        /// <exception cref="ArgumentNullException">If code is null.</exception>
        /// <exception cref="ArgumentException">If code is blank.</exception>
        public void AddCode(string code)
        {
            AddCode(code, null);
        }

        /// <summary>
        /// Adds the specified code to the scripting engine.
        /// The code that is added will be availible only under the namespace specified 
        /// instead of in the global scope of the script. 
        /// </summary>
        /// <param name="code">The code to be added.</param>
        /// <param name="namespaceName">The name of the namespace to add the code to.</param>
        /// <exception cref="ArgumentNullException">If code is null.</exception>
        /// <exception cref="ArgumentException">If code is blank.</exception>
        public void AddCode(string code, string namespaceName)
        {
            AddCode(code, namespaceName, null);
        }

        /// <summary>
        /// Adds the specified code to the scripting engine under the specified namespace.
        /// The code that is added will be availible only under the namespace specified 
        /// instead of in the global scope of the script. The script name will be used
        /// to provide more useful error information if an error occurs during the exection
        /// of the code that was added.
        /// </summary>
        /// <param name="code">The code to be added.</param>
        /// <param name="namespaceName">The name of the namespace to add the code to.</param>
        /// <param name="scriptName">The script name that the code came from.</param>
        /// <exception cref="ArgumentNullException">If code is null.</exception>
        /// <exception cref="ArgumentException">If code is blank.</exception>
        public void AddCode(string code, string namespaceName, string scriptName)
        {
            if (code == null)
            {
                throw new ArgumentNullException("code");
            }

            if (string.IsNullOrEmpty(code))
            {
                throw new ArgumentException("code parameter must contain code", "code");
            }

            if (namespaceName != null)
            {
                activeScript.AddNamedItem(namespaceName, ScriptItemFlags.CodeOnly | ScriptItemFlags.IsVisible);
            }

            try
            {
                /*
                 * In the event that the passed in script is not valid syntax
                 * an error will be thrown by the script engine. This will be 
                 * handled by the OnScriptError event before the exception
                 * is thrown here so we need to set this variable to use
                 * in the OnScriptError block to figure out the script name
                 * since it won't have been added to the script list.
                 */
                scriptToParse = scriptName;

                EXCEPINFO exceptionInfo = new EXCEPINFO();

                ulong cookie = (ulong)scripts.Count;

                parser.ParseScriptText(
                    code: code,
                    itemName: namespaceName,
                    context: null,
                    delimiter: null,
                    sourceContext: cookie,
                    startingLineNumber: 1u,
                    flags: ScriptTextFlags.IsVisible,
                    pVarResult: IntPtr.Zero,
                    excepInfo: out exceptionInfo);

                ScriptInfo si = new ScriptInfo()
                {
                    Code = code,
                    ScriptName = scriptName,
                    StartingLineNumber = 1u,
                    Cookie = cookie
                };

                scripts.Add(cookie, si);
            }
            finally
            {
                scriptToParse = null;
            }
        }

        /// <summary>
        /// Evaluates the specified code inside the context of the script and returns
        /// the result, or null if no result was returned.
        /// </summary>
        /// <param name="code">The code to evaluate.</param>
        /// <returns>The result of the evaluation.</returns>
        /// <exception cref="ArgumentNullException">If code is null.</exception>
        /// <exception cref="ArgumentException">If code is blank.</exception>
        public object Evaluate(string code)
        {
            if (code == null)
            {
                throw new ArgumentNullException("code");
            }

            if (string.IsNullOrEmpty(code))
            {
                throw new ArgumentException("code parameter must contain code", "code");
            }

            EXCEPINFO exceptionInfo = new EXCEPINFO();

            ulong cookie = (ulong)scripts.Count;

            IntPtr pResult = Marshal.AllocCoTaskMem(1024);

            try
            {
                parser.ParseScriptText(
                    code: code,
                    itemName: null,
                    context: null,
                    delimiter: null,
                    sourceContext: cookie,
                    startingLineNumber: 1u,
                    flags: ScriptTextFlags.IsExpression,
                    pVarResult: pResult,
                    excepInfo: out exceptionInfo);

                if (pResult != IntPtr.Zero)
                {
                    return Marshal.GetObjectForNativeVariant(pResult);
                }
            }
            finally
            {
                Marshal.FreeCoTaskMem(pResult);
            }

            return null;
        }

        /// <summary>
        /// Puts the scripting engine into the Started state.
        /// At this point code will be executed in the order they were added to the script engine.
        /// </summary>
        public void Start()
        {
            activeScript.SetScriptState(ScriptState.Started);
            activeScript.SetScriptState(ScriptState.Connected);
        }

        /// <summary>
        /// Stops and disposes this ActiveScriptEngine.
        /// </summary>
        public void Dispose()
        {
            if (activeScript != null)
            {
                try
                {
                    activeScript.Close();
                }
                finally
                {
                    activeScript = null;
                }
            }

            if (parser != null)
            {
                parser = null;
            }
        }

        /// <summary>
        /// Gets the IDispatch handle for the root namespace.
        /// </summary>
        /// <returns>An IDispatch handle for the root namespace.</returns>
        public object GetScriptHandle()
        {
            return GetScriptHandle(null);
        }

        /// <summary>
        /// Gets the IDispatch handle for the specified namespace.
        /// </summary>
        /// <param name="namespaceName">The namespace of the IDispatch handle to get, or null
        /// to get the root namespace.</param>
        /// <returns>An IDispatch handle.</returns>
        public object GetScriptHandle(string namespaceName)
        {
            return activeScript.GetScriptDispatch(namespaceName);
        }

        internal void OnScriptError(IActiveScriptError error)
        {
            this.LastError = new ScriptErrorInfo(error, this.scripts);

            if (this.LastError.ScriptName == null && scriptToParse != null)
            {
                this.LastError.ScriptName = scriptToParse;
            }

            var errorOccurred = this.ScriptErrorOccurred;

            if (errorOccurred != null)
            {
                errorOccurred(this, this.LastError);
            }
        }

        private T CreateInstanceFromProgID<T>(string progID) where T : class
        {
            Type type = Type.GetTypeFromProgID(progID);

            if (type != null)
            {
                object obj = Activator.CreateInstance(type);

                if (obj != null)
                {
                    return obj as T;
                }
            }

            return default(T);
        }
    }
}
