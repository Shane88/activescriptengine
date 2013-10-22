namespace ActiveXScriptLib
{
    using System;
    using Interop.ActiveXScript;
    using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

    internal class ActiveScriptParse
    {
        private IActiveScriptParse32 activeScriptParse32;
        private IActiveScriptParse64 activeScriptParse64;
        private bool is64Bit;

        public static ActiveScriptParse MakeActiveScriptParse(object activeScriptParse)
        {
            IActiveScriptParse64 parser64 = activeScriptParse as IActiveScriptParse64;

            if (parser64 != null)
            {
                return new ActiveScriptParse(parser64);
            }

            IActiveScriptParse32 parser32 = activeScriptParse as IActiveScriptParse32;

            if (parser32 != null)
            {
                return new ActiveScriptParse(parser32);
            }

            throw new ArgumentException("Unable to get parser interface", "activeScriptParse");
        }

        public ActiveScriptParse(IActiveScriptParse32 activeScriptParse)
        {
            if (activeScriptParse == null)
            {
                throw new ArgumentNullException("activeScriptParse");
            }

            this.activeScriptParse32 = activeScriptParse;
        }

        public ActiveScriptParse(IActiveScriptParse64 activeScriptParse)
        {
            if (activeScriptParse == null)
            {
                throw new ArgumentNullException("activeScriptParse");
            }

            this.activeScriptParse64 = activeScriptParse;

            this.is64Bit = true;
        }

        public void InitNew()
        {
            if (is64Bit)
            {
                activeScriptParse64.InitNew();
            }
            else
            {
                activeScriptParse32.InitNew();
            }
        }

        public void AddScriptlet(
            string defaultName,
            string code,
            string itemName,
            string subItemName,
            string eventName,
            string delimiter,
            ulong sourceContext,
            uint startingLineNumber,
            ScriptTextFlags flags,
            out string name,
            out EXCEPINFO excepInfo)
        {
            if (is64Bit)
            {
                activeScriptParse64.AddScriptlet(
                    defaultName,
                    code,
                    itemName,
                    subItemName,
                    eventName,
                    delimiter,
                    sourceContext,
                    startingLineNumber,
                    flags,
                    out name,
                    out excepInfo);
            }
            else
            {
                activeScriptParse32.AddScriptlet(
                    defaultName,
                    code,
                    itemName,
                    subItemName,
                    eventName,
                    delimiter,
                    (uint)sourceContext,
                    startingLineNumber,
                    flags,
                    out name,
                    out excepInfo);
            }
        }

        public void ParseScriptText(
            string code,
            string itemName,
            object context,
            string delimiter,
            ulong sourceContext,
            uint startingLineNumber,
            ScriptTextFlags flags,
            IntPtr pVarResult,
            out EXCEPINFO excepInfo)
        {
            if (is64Bit)
            {
                activeScriptParse64.ParseScriptText(
                    code,
                    itemName,
                    context,
                    delimiter,
                    sourceContext,
                    startingLineNumber,
                    flags,
                    pVarResult,
                    out excepInfo);
            }
            else
            {
                activeScriptParse32.ParseScriptText(
                    code,
                    itemName,
                    context,
                    delimiter,
                    (uint)sourceContext,
                    startingLineNumber,
                    flags,
                    pVarResult,
                    out excepInfo);
            }
        }
    }
}
