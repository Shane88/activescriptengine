namespace ActiveXScriptLib
{
    internal class ScriptInfo
    {
        public string ScriptName { get; set; }
        public string Code { get; set; }
        public ulong Cookie { get; set; }
        public uint StartingLineNumber { get; set; }
    }
}
