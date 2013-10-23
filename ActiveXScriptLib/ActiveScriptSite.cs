namespace ActiveXScriptLib
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    internal class ActiveScriptSite : IActiveScriptSite
    {
        private ActiveScriptEngine scriptEngine;

        public ActiveScriptSite(ActiveScriptEngine scriptEngine)
        {
            this.scriptEngine = scriptEngine;
        }

        public void GetLCID(out uint lcid)
        {
            lcid = 0;
        }

        public void GetItemInfo(string name, ScriptInfoFlags mask, ref IntPtr pUnkItem, ref IntPtr pTypeInfo)
        {
            if (mask == ScriptInfoFlags.IUnknown)
            {
                // Look up list of host objects.
                if (scriptEngine.hostObjects.ContainsKey(name))
                {
                    pUnkItem = Marshal.GetIUnknownForObject(scriptEngine.hostObjects[name]);
                }
                else
                {
                    object disp;
                    scriptEngine.activeScript.GetScriptDispatch(name, out disp);
                    pUnkItem = Marshal.GetIUnknownForObject(disp);
                }
            }
        }

        public void GetDocVersionString(out string version)
        {
            // TODO: What should this be?
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
            this.scriptEngine.OnScriptError(error);
        }

        public void OnEnterScript()
        {
        }

        public void OnLeaveScript()
        {
        }
    }
}
