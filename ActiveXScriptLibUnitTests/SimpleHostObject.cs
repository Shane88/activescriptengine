namespace ActiveXScriptLibUnitTests
{
    using System.Diagnostics;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    public class SimpleHostObject
    {
        public void Echo(string text)
        {
            Trace.WriteLine(text);
        }
    }
}
