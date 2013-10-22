namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;    

    public class ActiveScriptEngine : IDisposable
    {
        public event ScriptErrorOccurredDelegate ScriptErrorOccurred;
        public ScriptErrorInfo LastError { get; internal set; }

        private ActiveScriptParse parser;
        internal IActiveScript activeScript;
        internal Dictionary<string, object> hostObjects;
        private ActiveScriptSite scriptSite;

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

            hostObjects = new Dictionary<string, object>();
        }

        public void AddObject(string name, object obj)
        {
            hostObjects.Add(name, obj);
            activeScript.AddNamedItem(name, ScriptItemFlags.GlobalMembers | ScriptItemFlags.IsVisible);
        }

        public void AddCode(string code)
        {
            AddCode(code, null);
        }

        public void AddCode(string code, string namespaceName)
        {
            EXCEPINFO exceptionInfo = new EXCEPINFO();

            if (namespaceName != null)
            {
                activeScript.AddNamedItem(namespaceName, ScriptItemFlags.CodeOnly | ScriptItemFlags.IsVisible);
            }

            parser.ParseScriptText(
                code: code,
                itemName: namespaceName,
                context: null,
                delimiter: null,
                sourceContext: 0u,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsVisible,
                pVarResult: IntPtr.Zero,
                excepInfo: out exceptionInfo);
        }

        /// <summary>
        /// Puts the scripting engine into the initialized state.
        /// At this point all script text will be checked for syntax errors, but no code will be
        /// executed.
        /// </summary>
        public void Initialize()
        {
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

        public object GetScriptHandle()
        {
            return GetScriptHandle(null);
        }

        public object GetScriptHandle(string namespaceName)
        {
            object IDispatchInstance;
            activeScript.GetScriptDispatch(namespaceName, out IDispatchInstance);
            return IDispatchInstance;
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

        internal void OnScriptError(IActiveScriptError error)
        {
            this.LastError = new ScriptErrorInfo(error);

            var errorOccurred = this.ScriptErrorOccurred;

            if (errorOccurred != null)
            {
                errorOccurred(this, this.LastError);
            }
        }

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
    }
}
