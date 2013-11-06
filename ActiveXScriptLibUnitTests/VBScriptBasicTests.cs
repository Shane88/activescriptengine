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
            scriptEngine.Run();
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
            scriptEngine.Run();
        }

        [TestMethod]
        public void SimpleExecuteWithPublicFunction()
        {
            scriptEngine.AddCode(AddFunctionCode);
            scriptEngine.Run();

            dynamic script = scriptEngine.GetScriptHandle();
            int result = script.Add(1, 5);
            Assert.AreEqual(6, result);
        }

        [TestMethod]
        public void SimpleExecuteWithError()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146828277 // Divide By Zero Error.
            };

            scriptEngine.AddCode("a = 1 / 0");
            scriptEngine.Run();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827263, // Expected end of statement.
                ScriptName = "SimpleExecuteWithSyntaxError",
                LineNumber = 1
            };

            scriptEngine.AddCode("Dim a = Len(\"a", null, "SimpleExecuteWithSyntaxError");
            scriptEngine.Run();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError2()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827244, // Cannot use parentheses when calling a Sub.
                ScriptName = "SimpleExecuteWithSyntaxError2",
                LineNumber = 1
            };

            scriptEngine.AddCode("Thing(a, b)", null, "SimpleExecuteWithSyntaxError2");
            scriptEngine.Run();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError3()
        {
            this.expectedException = new ScriptErrorInfo()
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

            scriptEngine.AddCode(sb.ToString(), null, "SimpleExecuteWithSyntaxError3_1");
            scriptEngine.Run();

            scriptEngine.AddCode("Add 1, 2", null, "SimpleExecuteWithSyntaxError3_2");
        }
    }
}
