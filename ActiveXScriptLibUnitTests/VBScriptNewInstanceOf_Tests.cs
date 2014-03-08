namespace ActiveXScriptLibUnitTests
{
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