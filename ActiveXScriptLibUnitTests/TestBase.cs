namespace ActiveXScriptLibUnitTests
{
    using ActiveXScriptLib;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Diagnostics;

    public class TestBase
    {
        protected ActiveScriptEngine scriptEngine;
        protected ScriptErrorInfo expectedException;

        private string progID;

        public TestBase(string progID)
        {
            this.progID = progID;
        }

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
            }

            expectedException = null;
        }
    }

    public class VBScriptTestBase : TestBase
    {
        public VBScriptTestBase()
            : base("VBScript")
        {

        }
    }
}
