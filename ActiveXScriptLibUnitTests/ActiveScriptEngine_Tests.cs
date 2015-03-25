namespace ActiveXScriptLibUnitTests
{
   using System;
   using System.Runtime.InteropServices;
   using ActiveXScriptLib;
   using FluentAssertions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;
   using ActiveXScriptLib.Extensions;

   [TestClass]
   public class ActiveScriptEngine_Tests
   {
      [TestMethod]
      public void When_ActiveScriptEngine_is_passed_a_progId_which_is_not_a_ActiveScript_object()
      {
         // Arrange.
         // Act.
         Action action = () => new ActiveScriptEngine("Scripting.FileSystemObject");

         // Assert.
         action.ShouldThrow<InvalidActiveScriptClassException>();
      }

      [TestMethod]
      public void When_ActiveScriptEngine_is_passed_a_progId_which_is_not_valid()
      {
         // Arrange.
         // Act.
         Action action = () => new ActiveScriptEngine("ASKLDhjayd894ythfiujsd8fy9w");

         // Assert.
         action.ShouldThrow<COMException>();
      }

      [TestMethod]
      public void When_ActiveScriptEngine_is_passed_a_valid_ActiveScript_instance()
      {
         // Arrange.
         Type type = Type.GetTypeFromProgID("VBScript", true);         

         object activeScriptInstance = Activator.CreateInstance(type);

         // Act.
         ActiveScriptEngine engine = new ActiveScriptEngine(activeScriptInstance);

         // Assert.
         engine.Should().NotBeNull();
      }

      [TestMethod]
      public void When_ActiveScriptEngine_is_created_from_a_ActiveScript_instance_ProgId_should_be_null()
      {
         Type type = Type.GetTypeFromProgID("VBScript", true);         

         object activeScriptInstance = Activator.CreateInstance(type);

         // Act.
         ActiveScriptEngine engine = new ActiveScriptEngine(activeScriptInstance);

         // Assert.
         engine.ProgId.Should().Be(null);
      }

      [TestMethod]
      public void When_ActiveScriptEngine_is_connected_it_IsRunning()
      {
         // Arrange.
         ActiveScriptEngine engine = new ActiveScriptEngine(VBScript.ProgId);

         bool isRunningBeforeStart = engine.IsRunning;

         // Act.
         engine.Start();

         // Assert.
         isRunningBeforeStart.Should().BeFalse();
         engine.IsRunning.Should().BeTrue();
      }
   }
}