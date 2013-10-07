namespace ActiveXScriptLib
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;
    using System.Runtime.InteropServices;
    using System.Threading;

    public class ActiveScriptEngine : IActiveScriptSite
    {
        private IActiveScript activeScript;
        private IActiveScriptParse32 parser;

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
            
            parser = activeScript as IActiveScriptParse32;

            if (parser == null) 
            {
                throw new Exception("Unable to obtain IActiveScriptParse32 interface from script engine");
            }

            parser.InitNew();

            activeScript.SetScriptSite(this);

            activeScript.SetScriptState(ScriptState.Initialized);
        }

        public void AddCode(string code)
        {
            EXCEPINFO exceptionInfo = new EXCEPINFO();

            parser.ParseScriptText(
                code: code,
                itemName: null,
                context: null,
                delimiter: null,
                sourceContext: 0u,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsVisible,
                pVarResult: IntPtr.Zero,
                excepInfo: out exceptionInfo);
        }

        public void AddCode(string code, string alias)
        {
            EXCEPINFO exceptionInfo = new EXCEPINFO();

            activeScript.AddNamedItem(alias,ScriptItemFlags.IsVisible | ScriptItemFlags.GlobalMembers);

            parser.ParseScriptText(
                code: code,
                itemName: alias,
                context: null,
                delimiter: null,
                sourceContext: 0u,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsVisible,
                pVarResult: IntPtr.Zero,
                excepInfo: out exceptionInfo);
        }

        public void Run()
        {
            activeScript.SetScriptState(ScriptState.Started);
            activeScript.SetScriptState(ScriptState.Connected);
        }

        public object GetIDispatch()
        {
            object IDispatchInstance;
            activeScript.GetScriptDispatch(null, out IDispatchInstance);
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

        public void GetLCID(out uint lcid)
        {
            lcid = (uint)Thread.CurrentThread.CurrentUICulture.LCID;
        }

        public void GetItemInfo(string name, ScriptInfoFlags mask, ref IntPtr pUnkItem, ref IntPtr pTypeInfo)
        {
            //object disp;
            //activeScript.GetScriptDispatch(name, out disp);
            //pUnkItem = Marshal.GetIUnknownForObject(disp);
        }

        public void GetDocVersionString(out string version)
        {
            version = "1.0.0.0";
        }

        public void OnScriptTerminate(IntPtr pVarResult, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO excepInfo)
        {
        }

        public void OnStateChange(ScriptState state)
        {
        }

        public void OnScriptError(IActiveScriptError error)
        {
        }

        public void OnEnterScript()
        {
        }

        public void OnLeaveScript()
        {
        }
    }
}
