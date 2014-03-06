namespace ActiveXScriptLib
{
   using System;
   using System.Runtime.InteropServices;
   using Interop.ActiveXScript;
   using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    internal class ActiveScriptSite : IActiveScriptSite
    {
        private ActiveScriptEngine scriptEngine;

        public ActiveScriptSite(ActiveScriptEngine scriptEngine)
        {
            this.scriptEngine = scriptEngine;
        }

        public uint GetLCID()
        {
            return 0;
        }

        public void GetItemInfo(string name, ScriptInfoFlags mask, ref IntPtr pUnkItem, ref IntPtr pTypeInfo)
        {
            if (mask == ScriptInfoFlags.IUnknown)
            {
                // Look up list of host objects.
                if (scriptEngine.HostObjects.ContainsKey(name))
                {
                    pUnkItem = Marshal.GetIUnknownForObject(scriptEngine.HostObjects[name]);
                }
                else
                {
                    object disp = scriptEngine.ActiveScript.GetScriptDispatch(name);
                    pUnkItem = Marshal.GetIUnknownForObject(disp);
                }
            }
        }

        public string GetDocVersionString()
        {
            throw new NotImplementedException();
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
