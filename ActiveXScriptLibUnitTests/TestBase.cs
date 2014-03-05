namespace ActiveXScriptLibUnitTests
{
   using ActiveXScriptLib;
   using Microsoft.VisualStudio.TestTools.UnitTesting;
   using System;
   using System.Diagnostics;

   public class TestBase
   {
      protected ActiveScriptEngine scriptEngine;
      protected ScriptErrorInfo expectedError;
      protected SimpleHostObject WScript;
      protected dynamic script;

      private string progID;

      public TestBase(string progID)
      {
         this.progID = progID;
      }

      [TestInitialize]
      public void GlobalTestSetup()
      {
         Trace.WriteLine("Environment is " + (Environment.Is64BitProcess ? "64bit" : "32bit"));

         this.WScript = new SimpleHostObject();
         this.scriptEngine = new ActiveScriptEngine(progID);
         this.scriptEngine.ScriptErrorOccurred += scriptEngine_ScriptErrorOccurred;
         this.scriptEngine.AddObject("WScript", this.WScript);
         this.script = this.scriptEngine.GetScriptHandle();

         this.expectedError = null;
      }

      [TestCleanup]
      public void GlobalTestCleanup()
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

         if (expectedError == null)
         {
            Assert.Fail("Script threw an unexpected error");
         }
         else
         {
            CheckErrors(expectedError, error);
         }
      }

      public static void CheckErrors(ScriptErrorInfo expected, ScriptErrorInfo actual)
      {
         Assert.AreEqual(expected.ErrorNumber, actual.ErrorNumber);

         if (expected.ColumnNumber != 0)
         {
            Assert.AreEqual(expected.ColumnNumber, actual.ColumnNumber);
         }

         if (expected.LineNumber != 0)
         {
            Assert.AreEqual(expected.LineNumber, actual.LineNumber);
         }

         if (!string.IsNullOrEmpty(expected.ScriptName))
         {
            Assert.AreEqual(expected.ScriptName, actual.ScriptName);
         }
      }
   }

   public class VBScriptTestBase : TestBase
   {
      public const string AddFunctionCode
          = "Public Function Add(a, b) " +
            "   Add = a + b " +
            "End Function";

      public VBScriptTestBase()
         : base(VBScript.ProgID)
      {

      }
   }

   public class JScriptTestBase : TestBase
   {
      public const string AddFunctionCode
          = "function Add(a, b) { " +
            "   return a + b; " +
            "}";

      public JScriptTestBase()
         : base(JScript.ProgID)
      {

      }
   }
}
