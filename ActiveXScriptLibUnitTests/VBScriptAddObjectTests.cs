namespace ActiveXScriptLibUnitTests
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Runtime.InteropServices;
    using System.Diagnostics;

    [TestClass]
    public class VBScriptAddObjectTests : VBScriptTestBase
    {
        [TestMethod]
        public void AddBasicObject()
        {
            BasicMathObject bmo = new BasicMathObject()
            {
                ExpectedA = 1,
                ExpectedB = 5
            };

            scriptEngine.AddObject("Math", bmo);

            scriptEngine.AddCode("Math.Add 1, 5");

            scriptEngine.Start();

            bmo.ExpectedA = 2;
            bmo.ExpectedB = 7;

            scriptEngine.AddCode("Add 2, 7"); // TODO: Why does this work?

            bmo.ExpectedA = 23;
            bmo.ExpectedB = 57;

            dynamic script = scriptEngine.GetScriptHandle();

            Assert.AreEqual(80, script.Math.Add(23, 57));
        }
    }

    [ComVisible(true)]
    public class BasicMathObject
    {
        public int ExpectedA { get; set; }
        public int ExpectedB { get; set; }

        public int Add(int a, int b)
        {
            Trace.WriteLine("ExpectedA=" + ExpectedA + " Actual=" + a);
            Trace.WriteLine("ExpectedA=" + ExpectedB + " Actual=" + b);

            Assert.AreEqual(ExpectedA, a);
            Assert.AreEqual(ExpectedB, b);

            return a + b;
        }
    }
}
