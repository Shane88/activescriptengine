namespace ActiveXScriptLibUnitTests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class VBScriptNamespaceTests : VBScriptTestBase
    {
        [TestMethod]
        public void AddNamespaceTest()
        {
            WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("3", text);
            };

            scriptEngine.AddCode(AddFunctionCode, "Math");

            scriptEngine.AddCode("WScript.Echo Math.Add(1,2)");

            scriptEngine.Run();

            dynamic script = scriptEngine.GetScriptHandle();

            Assert.AreEqual(7, script.Math.Add(3, 4));
        }

        [TestMethod]
        public void CallGlobalFunctionsFromNamespace()
        {
            WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("102", text);
            };

            scriptEngine.AddCode(AddFunctionCode);

            scriptEngine.AddCode("Public Sub DoAdd() WScript.Echo Add(100,2) End Sub");

            scriptEngine.Run();

            dynamic script = scriptEngine.GetScriptHandle();

            Assert.AreEqual(7, script.Add(3, 4));
        }

        [TestMethod]
        public void NamespaceCallIntoNamespace()
        {
            WScript.OnEcho += (text) =>
            {
                Assert.AreEqual("103", text);
            };

            scriptEngine.AddCode(AddFunctionCode, "Math");

            scriptEngine.AddCode("Public Sub DoAdd() WScript.Echo Math.Add(98,5) End Sub", "Worker");

            scriptEngine.AddCode("Worker.DoAdd");

            scriptEngine.Run();

            dynamic script = scriptEngine.GetScriptHandle();

            script.Worker.DoAdd();

            script = scriptEngine.GetScriptHandle("Worker");
            script.DoAdd();
        }
    }
}
