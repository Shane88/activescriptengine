namespace ActiveXScriptLibUnitTests
{
    using ActiveXScriptLib;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Diagnostics;

    public class TestBase
    {
        protected ActiveScriptEngine scriptEngine;
        protected ScriptErrorInfo expectedException;
        protected SimpleHostObject WScript;

        private string progID;

        public TestBase(string progID)
        {
            this.progID = progID;
        }

        [TestInitialize]
        public void TestSetup()
        {
            Trace.WriteLine("Environment is " + (Environment.Is64BitProcess ? "64bit" : "32bit"));

            this.WScript = new SimpleHostObject();
            this.scriptEngine = new ActiveScriptEngine(VBScript.ProgID);
            this.scriptEngine.ScriptErrorOccurred += scriptEngine_ScriptErrorOccurred;
            this.scriptEngine.AddObject("WScript", this.WScript);
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

        protected void scriptEngine_ScriptErrorOccurred(ActiveScriptEngine sender, ScriptErrorInfo error)
        {
            Trace.WriteLine(error.DebugDump());

            if (expectedException == null)
            {
                Assert.Fail("Script threw an unexpected error");
            }
            else
            {
                Assert.AreEqual(expectedException.ErrorNumber, error.ErrorNumber);

                if (expectedException.ColumnNumber != 0)
                {
                    Assert.AreEqual(expectedException.ColumnNumber, error.ColumnNumber);
                }

                if (expectedException.LineNumber != 0)
                {
                    Assert.AreEqual(expectedException.LineNumber, error.LineNumber);
                }

                if (!string.IsNullOrEmpty(expectedException.ScriptName))
                {
                    Assert.AreEqual(expectedException.ScriptName, error.ScriptName);
                }
            }

            expectedException = null;
        }
    }

    public class VBScriptTestBase : TestBase
    {
        public const string AddFunctionCode
            = "Public Function Add(a, b) " +
              "   Add = a + b " +
              "End Function";

        public VBScriptTestBase()
            : base("VBScript")
        {

        }
    }
}
