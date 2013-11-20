namespace ActiveXScriptLibUnitTests
{
    using ActiveXScriptLib;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Runtime.InteropServices;
    using System.Text;

    [TestClass]
    public class VBScriptBasicTests : VBScriptTestBase
    {
        [TestMethod]
        public void SimpleExecute()
        {
            this.WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("Echo from VBScript", text);
            };

            scriptEngine.AddCode("WScript.Echo \"Echo from VBScript\"");
            scriptEngine.Start();
        }

        [TestMethod]
        public void TestEval()
        {
            scriptEngine.Start();

            Assert.AreEqual((short)2, scriptEngine.Eval("1 + 1"));
        }

        [TestMethod]
        public void SimpleExecuteWithFunction()
        {
            this.WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("4", text);
            };

            scriptEngine.AddCode(AddFunctionCode);
            scriptEngine.AddCode("WScript.Echo Add(1, 3)");
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithPublicFunction()
        {
            scriptEngine.AddCode(AddFunctionCode);
            scriptEngine.Start();

            dynamic script = scriptEngine.GetScriptHandle();
            int result = script.Add(1, 5);
            Assert.AreEqual(6, result);
        }

        [TestMethod]
        public void SimpleExecuteWithError()
        {
            this.expectedError = new ScriptErrorInfo()
            {
                ErrorNumber = -2146828277 // Divide By Zero Error.
            };

            scriptEngine.AddCode("a = 1 / 0");
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithSyntaxError()
        {
            this.expectedError = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827263, // Expected end of statement.
                ScriptName = "SimpleExecuteWithSyntaxError",
                LineNumber = 1
            };

            try
            {
                scriptEngine.AddCode("Dim a = Len(\"a", null, "SimpleExecuteWithSyntaxError");
                scriptEngine.Start();
            }
            catch (COMException)
            {
                CheckErrors(expectedError, scriptEngine.LastError);
            }
        }

        [TestMethod]
        public void SimpleExecuteWithSyntaxError2()
        {
            this.expectedError = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827244, // Cannot use parentheses when calling a Sub.
                ScriptName = "SimpleExecuteWithSyntaxError2",
                LineNumber = 1
            };

            try
            {
                scriptEngine.AddCode("Thing(a, b)", null, "SimpleExecuteWithSyntaxError2");
                scriptEngine.Start();
            }
            catch (COMException)
            {
                CheckErrors(expectedError, scriptEngine.LastError);
            }
        }

        [TestMethod]
        public void SimpleExecuteWithSyntaxError3()
        {
            this.expectedError = new ScriptErrorInfo()
            {
                ErrorNumber = -2146828277, // Divide By Zero Error.
                ScriptName = "SimpleExecuteWithSyntaxError3_1",
                LineNumber = 2
            };

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Public Function Add(a, b)");
            sb.AppendLine("   Add = (a + b) / 0");
            sb.AppendLine("End Function");
            sb.AppendLine();

            try
            {
                scriptEngine.AddCode(sb.ToString(), null, "SimpleExecuteWithSyntaxError3_1");
                scriptEngine.Start();

                scriptEngine.AddCode("Add 1, 2", null, "SimpleExecuteWithSyntaxError3_2");
            }
            catch (COMException)
            {
                CheckErrors(expectedError, scriptEngine.LastError);
            }
        }
    }
}
