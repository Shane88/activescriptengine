namespace ActiveXScriptLibUnitTests
{
    using ActiveXScriptLib;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;

    [TestClass]
    public class VBScriptBasicTests : VBScriptTestBase
    {
        [TestMethod]
        public void SimpleExecute()
        {
            scriptEngine.AddCode("WScript.Echo \"Echo from VBScript\"");

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
            scriptEngine.Start();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827263 // Expected end of statement.
            };

            scriptEngine.AddCode("Dim a = Len(\"a", null, "SimpleExecuteWithSyntaxError");
            scriptEngine.Start();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError2()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146827244 // Cannot use parentheses when calling a Sub.
            };

            scriptEngine.AddCode("Thing(a, b)", null, "SimpleExecuteWithSyntaxError2");
            scriptEngine.Start();
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void SimpleExecuteWithSyntaxError3()
        {
            this.expectedException = new ScriptErrorInfo()
            {
                ErrorNumber = -2146828277 // Divide By Zero Error.
            };

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Public Function Add(a, b)");
            sb.AppendLine("   Add = (a + b) / 0");
            sb.AppendLine("End Function");
            sb.AppendLine();

            scriptEngine.AddCode(sb.ToString(), null, "SimpleExecuteWithSyntaxError3_1");
            scriptEngine.Start();

            scriptEngine.AddCode("Add 1, 2", null, "SimpleExecuteWithSyntaxError3_2");
        }
    }
}
