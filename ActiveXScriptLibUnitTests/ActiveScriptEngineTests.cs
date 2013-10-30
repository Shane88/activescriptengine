namespace ActiveXScriptLibUnitTests
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using ActiveXScriptLib;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;

    [TestClass]
    public class ActiveScriptEngineTests
    {
        private ActiveScriptEngine scriptEngine;
        private ScriptErrorInfo expectedException;

        [TestInitialize]
        public void TestSetup()
        {
            this.scriptEngine = new ActiveScriptEngine(VBScript.ProgID);
            this.scriptEngine.ScriptErrorOccurred += scriptEngine_ScriptErrorOccurred;
            this.scriptEngine.AddObject("WScript", new SimpleHostObject());
        }

        [TestCleanup]
        public void TestCleanup()
        {
            if (this.scriptEngine != null)
            {
                this.scriptEngine.ScriptErrorOccurred -= scriptEngine_ScriptErrorOccurred;
                this.scriptEngine.Dispose();
            }
        }

        private void scriptEngine_ScriptErrorOccurred(ActiveScriptEngine sender, ScriptErrorInfo error)
        {
            Trace.WriteLine(error.DebugDump());

            if (expectedException == null)
            {
                Assert.Fail("Script threw an unexpected error");
            }
            else
            {
                Assert.AreEqual(expectedException.ErrorNumber, error.ErrorNumber);
            }
        }

        [TestMethod]
        public void SimpleExecute()
        {
            scriptEngine.AddCode("WScript.Echo \"Echo from VBScript\"");

            scriptEngine.Initialize();
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithFunction()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Private Function Add(a, b)");
            sb.AppendLine("   Add = a + b");
            sb.AppendLine("End Function");
            sb.AppendLine();
            sb.AppendLine("WScript.Echo \"Result of Add = \" & Add(1, 3)");

            scriptEngine.AddCode(sb.ToString());

            scriptEngine.Initialize();
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithPublicFunction()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Public Function Add(a, b)");
            sb.AppendLine("   Add = a + b");
            sb.AppendLine("End Function");
            sb.AppendLine();

            scriptEngine.AddCode(sb.ToString());

            scriptEngine.Initialize();
            scriptEngine.Start();

            dynamic script = scriptEngine.GetScriptHandle();
            int result = script.Add(1, 5);
            Trace.WriteLine("script.Add(1, 5) = " + result);
        }

        [TestMethod]
        public void SimpleExecuteWithError()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146828277 // Divide By Zero Error.
            };

            scriptEngine.AddCode("a = 1 / 0");

            scriptEngine.Initialize();
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithSyntaxError()
        {
            scriptEngine.AddCode("Dim .a = 1 / 0");

            scriptEngine.Initialize();
            scriptEngine.Start();
        }
    }
}
