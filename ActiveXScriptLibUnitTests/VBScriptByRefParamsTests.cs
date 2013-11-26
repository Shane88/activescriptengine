namespace ActiveXScriptLibUnitTests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Runtime.ExceptionServices;
    using System.Runtime.InteropServices;

    /*
     * Known issues
     * 
     * To pass Nothing to VBScript you will need to use [return: MarshalAs(UnmanagedType.IDispatch)] and return null.
     * 
     * You will not be able to cast function pointers obtained from GetRef into managed delegate types.
     * 
     * Ref parameters will not be passed in and out properly when using a dynamic script. Out works fine.
     * 
     * 
     */
    [TestClass]
    public class VBScriptByRefParamsTests : VBScriptTestBase
    {
        [TestMethod]
        [HandleProcessCorruptedStateExceptions]
        public void DelegateTest()
        {
            scriptEngine.AddCode("Public Function GetTheRef(sFuncName) Set GetTheRef = GetRef(sFuncName) End Function");

            scriptEngine.AddCode(AddFunctionCode);

            scriptEngine.Start();

            dynamic script = scriptEngine.GetScriptHandle();

            object ptrAddFunc = script.GetTheRef("Add");

            IntPtr ptrFunc = Marshal.GetIUnknownForObject(ptrAddFunc);

            Delegate d = Marshal.GetDelegateForFunctionPointer(ptrFunc, typeof(AddDelegate));
            
            AddDelegate addDel = (AddDelegate)d;

            try
            {
                Assert.AreEqual(3, addDel(1, 2));
                Assert.Fail("Expected access violation exception");
            }
            catch (AccessViolationException)
            {
            }
        }

        public delegate int AddDelegate(int a, int b);
    }

}
