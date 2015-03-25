namespace ActiveXScriptLibUnitTests
{
   using System.Diagnostics;
   using System.Runtime.InteropServices;
   using ActiveXScriptLib.Extensions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class ActiveScriptEngineBuilder_Tests
   {
      [TestMethod]
      public void When_an_engine_is_built_with_stuff()
      {
         HostObject hostObject = new HostObject();

         var scriptEngine = new ActiveScriptEngineBuilder()
            .LogErrorsToTrace()
            .AddObject("WScript", hostObject)
            .AddObject<TestObject>("Test",
               (testObject) =>
               {
                  testObject.Name = "Hello World!";
               })
            .AddCode("Public Sub Print(text) \n WScript.Echo text \n End Sub")
            .AddCodeFile(@"..\..\MyAwesomeCode.vbs")
            .AddCreateObjectHook(
               factory =>
               {
                  factory.AddHook("gHost", hostObject);
                  factory.AddHook("Blah", hostObject);
                  factory.AddHook("blah2", hostObject);
               })
            .StartEngineOnBuild()
            .Build();

         dynamic script = scriptEngine.GetScriptHandle();

         script.Print("Hello World");
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
