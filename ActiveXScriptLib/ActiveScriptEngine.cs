namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    internal class ScriptInfo
    {
        public string ScriptName { get; set; }
        public string Code { get; set; }
        public uint Cookie { get; set; }
        public uint StartingLineNumber { get; set; }
    }

    public class ActiveScriptEngine : IDisposable
    {
        public event ScriptErrorOccurredDelegate ScriptErrorOccurred;
        public ScriptErrorInfo LastError { get; internal set; }

        internal IActiveScript activeScript;
        internal Dictionary<string, object> hostObjects;

        private ActiveScriptParse parser;
        private ActiveScriptSite scriptSite;

        private Dictionary<uint, ScriptInfo> scripts;

        public ActiveScriptEngine(string progID)
        {
            if (string.IsNullOrWhiteSpace(progID))
            {
                throw new ArgumentException("Argument must not be null or blank");
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
            //activeScript.SetScriptState(ScriptState.Uninitialized);

            hostObjects = new Dictionary<string, object>();
            scripts = new Dictionary<uint, ScriptInfo>();
        }

        public void AddObject(string name, object obj)
        {
            // TODO: Remove duplicate entries.
            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.GlobalMembers | ScriptItemFlags.IsVisible);
        }

        public void AddCode(string code)
        {
            AddCode(code, null);
        }

        public void AddCode(string code, string namespaceName)
        {
            AddCode(code, namespaceName, null);
        }

        /// <summary>
        /// Adds the specified code to the scripting engine.
        /// </summary>
        /// <param name="code"></param>
        /// <param name="namespaceName"></param>
        /// <param name="scriptName"></param>
        public void AddCode(string code, string namespaceName, string scriptName)
        {
            ScriptState ss = activeScript.GetScriptState();

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
                // TODO: What if theres already a named item by this name?
                activeScript.AddNamedItem(namespaceName, ScriptItemFlags.CodeOnly | ScriptItemFlags.IsVisible);
            }

            uint cookie = (uint)scripts.Count;

            // TODO: If this lines fails because of syntax error then the script name on the error info
            // will not be populated because we haven't stored the cookie for it yet.

            EXCEPINFO exceptionInfo = new EXCEPINFO();

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

        /// <summary>
        /// Puts the scripting engine into the initialized state.
        /// At this point all script text will be checked for syntax errors, but no code will be
        /// executed.
        /// </summary>
        public void Initialize()
        {
            // TODO: Syntax checking appears to be happening before this is called
            activeScript.SetScriptState(ScriptState.Initialized);
        }

        /// <summary>
        /// Puts the scripting engine into the Started state.
        /// At this point any code not within functions, subs, or classes will be executed.
        /// Code will be executed in the order they were added to the script engine.
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
                catch
                {
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
            object IDispatchInstance;
            activeScript.GetScriptDispatch(namespaceName, out IDispatchInstance);
            return IDispatchInstance;
        }

        internal void OnScriptError(IActiveScriptError error)
        {
            this.LastError = new ScriptErrorInfo(this.scripts, error);

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
