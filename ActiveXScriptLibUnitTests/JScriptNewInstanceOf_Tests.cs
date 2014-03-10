namespace ActiveXScriptLibUnitTests
{
   using ActiveXScriptLib.Extensions;
   using FluentAssertions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class JScriptNewInstanceOf_Tests : JScriptTestBase
   {
      [TestInitialize]
      public void TestSetup()
      {
         scriptEngine.AddCode(@"
            // Define a class like this
            function Person(){
               // Add object properties like this
               this.Name = '';
            }");

         scriptEngine.Start();
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