namespace ActiveXScriptLibUnitTests
{
   using System;
   using ActiveXScriptLib;
   using ActiveXScriptLib.Extensions;
   using FluentAssertions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class VBScriptNewInstanceOf_Tests : VBScriptTestBase
   {
      [TestInitialize]
      public void TestSetup()
      {
         scriptEngine.AddCode(
            @"Class Person
                 Public Name
                 Public DateOfBirth
              End Class");

         scriptEngine.Start();
      }

      [TestMethod]
      public void When_NewInstanceOf_is_called_with_a_null_engine()
      {
         // Arrange.
         ActiveScriptEngine engine = null;

         // Act.
         Action action = () => engine.NewInstanceOf("Person");

         // Assert.
         action.ShouldThrow<ArgumentNullException>();
      }

      [TestMethod]
      public void When_I_create_an_instance_of_a_class()
      {
         // Arrange.
         // Act.
         dynamic person = scriptEngine.NewInstanceOf("Person");

         person.Name = "Name set from C#";

         // Assert.
         string name = person.Name;
         name.Should().Be("Name set from C#");
      }
   }
}