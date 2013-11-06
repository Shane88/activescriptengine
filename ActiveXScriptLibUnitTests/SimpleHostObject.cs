namespace ActiveXScriptLibUnitTests
{
   using System;
   using System.Diagnostics;
   using System.Runtime.InteropServices;

    [ComVisible(true)]
    public class SimpleHostObject
    {
        public event Action<string> OnEcho;

        public void Echo(string text)
        {
            Trace.WriteLine(text);

            var onEcho = this.OnEcho;

            if (onEcho != null)
            {
                onEcho(text);
            }
        }
    }
}
