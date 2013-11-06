namespace ActiveXScriptLib
{
    using Interop.ActiveXScript;
    using System;
    using System.Collections.Generic;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;
    
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

        public void AddObject(string name, object obj)
        {
            // TODO: Remove duplicate entries.
            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible);
        }

        public void AddGlobalMemberObject(string name, object obj)
        {
            // TODO: Remove duplicate entries.
            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible | ScriptItemFlags.GlobalMembers);
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
        /// Puts the scripting engine into the Started state.
        /// At this point code will be executed in the order they were added to the script engine.
        /// </summary>
        public void Run()
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
