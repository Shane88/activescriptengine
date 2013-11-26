namespace ActiveXScriptLibUnitTests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;

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
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void CallObjectWithNoAlias()
        {
            BasicMathObject bmo = new BasicMathObject()
            {
                ExpectedA = 1,
                ExpectedB = 5
            };

            scriptEngine.AddObject("Math", bmo);

            scriptEngine.AddCode("Add 1, 5");

            scriptEngine.Start();

            bmo.ExpectedA = 23;
            bmo.ExpectedB = 57;
        }

        [TestMethod]
        public void AddBasicGlobalMemberObject()
        {
            BasicMathObject bmo = new BasicMathObject()
            {
                ExpectedA = 2,
                ExpectedB = 7
            };

            scriptEngine.AddGlobalMemberObject("Math", bmo);
            scriptEngine.Start();

            // No Math alias used.
            scriptEngine.AddCode("Add 2, 7");

            dynamic script = scriptEngine.GetScriptHandle();
        }

        [TestMethod]
        public void AddObjectsWithSameName()
        {
            BasicMathObject bmo = new BasicMathObject()
            {
                ExpectedA = 2,
                ExpectedB = 7
            };

            scriptEngine.AddObject("Math", bmo);

            try
            {
                scriptEngine.AddObject("Math", bmo);
                Assert.Fail("Expected ArgumentException");
            }
            catch (ArgumentException)
            {
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
                Trace.WriteLine("ExpectedB=" + ExpectedB + " Actual=" + b);

                Assert.AreEqual(ExpectedA, a);
                Assert.AreEqual(ExpectedB, b);

                return a + b;
            }
        }
    }
}
