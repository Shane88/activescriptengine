namespace ActiveXScriptLibUnitTests
{
    using ActiveXScriptLib;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    [TestClass]
    public class JScriptBasicTests : JScriptTestBase
    {

        [TestMethod]
        public void SimpleExecute()
        {
            this.WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("Echo from JScript", text);
            };

            scriptEngine.AddCode("WScript.Echo(\"Echo from JScript\");");
            scriptEngine.Start();
        }

        [TestMethod]
        public void TestEval()
        {
            scriptEngine.Start();

            Assert.AreEqual(2, scriptEngine.Evaluate("1 + 1"));
        }

        [TestMethod]
        public void SimpleExecuteWithFunction()
        {
            this.WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("4", text);
            };

            scriptEngine.AddCode(AddFunctionCode);
            scriptEngine.AddCode("WScript.Echo(Add(1, 3))");
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

            scriptEngine.AddCode("var a = 1 / 0");
            scriptEngine.Start();
        }

        [TestMethod]
        public void SimpleExecuteWithSyntaxError()
        {
            this.expectedError = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827273, // Unterminated string constant
                ScriptName = "SimpleExecuteWithSyntaxError",
                LineNumber = 1
            };

            try
            {
                scriptEngine.AddCode("var a = Len(\"a", null, "SimpleExecuteWithSyntaxError");
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
            sb.AppendLine("function Add(a, b) { ");
            sb.AppendLine("   return (a + b) / 0 ");
            sb.AppendLine("}");
            sb.AppendLine();

            try
            {
                scriptEngine.AddCode(sb.ToString(), null, "SimpleExecuteWithSyntaxError3_1");
                scriptEngine.Start();

                scriptEngine.AddCode("Add(1, 2)", null, "SimpleExecuteWithSyntaxError3_2");
            }
            catch (COMException)
            {
                CheckErrors(expectedError, scriptEngine.LastError);
            }
        }

        [TestMethod]
        public void TestNulls()
        {
            scriptEngine.AddCode("function IsNull(value) { return value == null; }");

            scriptEngine.AddCode("function IsUndefined(value) { return typeof value == \"undefined\"; }");

            scriptEngine.Start();

            dynamic script = scriptEngine.GetScriptHandle();

            Assert.IsTrue(script.IsNull(null));

            Assert.IsTrue(script.IsNull(JScript.Null));

            Assert.IsTrue(script.IsUndefined(JScript.Undefined));
        }
    }
}
