namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    public class ActiveScriptEngine : IActiveScriptSite, IDisposable
    {
        private IActiveScript activeScript;
        private IActiveScriptParse parser;
        private Dictionary<string, object> hostObjects;

        public ScriptExceptionInfo LastError { get; private set; }

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
            
            parser = activeScript as IActiveScriptParse;

            if (parser == null) 
            {
                throw new Exception("Unable to obtain IActiveScriptParse32 interface from script engine");
            }

            parser.InitNew();

            activeScript.SetScriptSite(this);

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
                pVarResult: null,
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

        /// <summary>
        /// TODO: Not working.
        /// </summary>
        /// <param name="expression"></param>
        /// <returns></returns>
        public object Eval(string expression)
        {
            EXCEPINFO exceptionInfo;
            object pResult = new object();

            parser.ParseScriptText(
                code: expression,
                itemName: null,
                context: null,
                delimiter: null,
                sourceContext: 0u,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsExpression,
                pVarResult: pResult,
                excepInfo: out exceptionInfo);

            return pResult;
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

        #region IActiveScriptSite Interface

        public void GetLCID(out uint lcid)
        {
            // TODO: What should we do here?
            lcid = (uint)Thread.CurrentThread.CurrentUICulture.LCID;
        }

        public void GetItemInfo(string name, ScriptInfoFlags mask, ref IntPtr pUnkItem, ref IntPtr pTypeInfo)
        {
            if (mask == ScriptInfoFlags.IUnknown)
            {
                // Look up list of host objects.
                if (hostObjects.ContainsKey(name))
                {
                    pUnkItem = Marshal.GetIUnknownForObject(hostObjects[name]);
                }
                else
                {
                    object disp;
                    activeScript.GetScriptDispatch(name, out disp);
                    pUnkItem = Marshal.GetIUnknownForObject(disp);
                }
            }
        }

        public void GetDocVersionString(out string version)
        {
            version = "1.0.0.0";
        }

        public void OnScriptTerminate(IntPtr pVarResult, ref EXCEPINFO excepInfo)
        {
        }

        public void OnStateChange(ScriptState state)
        {
        }

        public void OnScriptError(IActiveScriptError error)
        {
            this.LastError = new ScriptExceptionInfo(error);
        }

        public void OnEnterScript()
        {
        }

        public void OnLeaveScript()
        {
        }

        #endregion IActiveScriptSite Interface

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
