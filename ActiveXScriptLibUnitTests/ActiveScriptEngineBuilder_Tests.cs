namespace ActiveXScriptLibUnitTests
{
   using System.Diagnostics;
   using System.Runtime.InteropServices;
   using ActiveXScriptLib;
   using ActiveXScriptLib.Extensions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class ActiveScriptEngineBuilder_Tests
   {
      [TestMethod]
      public void When_an_engine_is_built_with_stuff()
      {
         HostObject hostObject = new HostObject();

         ActiveScriptEngine scriptEngine = new ActiveScriptEngineBuilder()

            // Log all errors from the script engine into Trace output.
            .LogErrorsToTrace()

            // Add any objects our scripts need here.
            .AddObject("WScript", hostObject)
            .AddObject<TestObject>("Test",
               (testObject) =>
               {
                  testObject.Name = "Hello World!";
               })

            // Add any code we need here.
            .AddCode("Public Sub Print(text) \n WScript.Echo text \n End Sub")
            .AddCodeFiles(@"..\..\*.vbs") // Wild cards can be used here.

            // Intercept CreateObject calls from the script and instead return the objects defined here.
            .InterceptCreateObject(factory => factory
                  .Intercept("gHost").With(hostObject)
                  .Intercept("Blah").With(hostObject)
                  .Intercept("blah2").With(hostObject))
            
            // Specify that we want the engine to be started as soon as we build.
            .StartEngineOnBuild()

            // Configure these actions to be run when the engine is started.
            // Actions will be run in the order specified.
            .OnStart(script => script.WScript.Echo("Echo from OnStart"))
            .RunCodeOnStart("WScript.Echo \"Echo from RunCodeOnStart\"")

            .Build();

         dynamic script2 = scriptEngine.GetScriptHandle();

         script2.Print("Echo from after engine is built");
      }

      [ComVisible(true)]
      public class TestObject
      {
         public string Name { get; set; }
      }

      [ComVisible(true)]
      public class HostObject
      {
         public void Echo(string text)
         {
            Trace.WriteLine(text);
         }
      }
   }
}
