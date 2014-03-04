namespace ActiveXScriptLibUnitTests
{
   using System;
   using ActiveXScriptLib.Extensions;
   using FluentAssertions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class FunctionOverrider_Tests : VBScriptTestBase
   {
      [TestMethod]
      public void When_I_override_a_simple_function()
      {
         // Arrange.
         string actualText = null;

         scriptEngine.AddCode(
            @"Public Sub Echo(sText)
                 ' Does nothing
              End Sub");

         scriptEngine.OverrideGlobalFunction<string>("Echo", (text) => actualText = text);

         // Act.
         script.Echo("Hello World from C#");

         // Assert.
         actualText.Should().Be("Hello World from C#");
      }

      [TestMethod]
      public void When_I_override_a_complex_function()
      {
         // Arrange.
         const int ExpectedNum = 5;
         const string ExpectedString = "Hello World";
         const bool ExpectedBool = true;
         DateTime ExpectedDateTime = DateTime.Now.Subtract(2.Days());
         object ExpectedObject = new object();

         int actualNum = 0;
         string actualString = null;
         bool actualBool = false;
         DateTime actualDate = new DateTime();
         object actualObject = null;

         scriptEngine.AddCode(
            @"Public Sub ComplexFunc(nNum, sText, bBool, dDate, oObject)
                 ' Does lots of stuff
                 Err.Raise 1
              End Sub");

         scriptEngine.OverrideGlobalFunction<int, string, bool, DateTime, object>(
            "ComplexFunc",
            (num, text, theBool, date, obj) =>
            {
               actualNum = num;
               actualString = text;
               actualBool = theBool;
               actualDate = date;
               actualObject = obj;
            });

         // Act.
         script.ComplexFunc(ExpectedNum, ExpectedString, ExpectedBool, ExpectedDateTime, ExpectedObject);

         // Assert.
         actualNum.Should().Be(ExpectedNum);
         actualString.Should().Be(ExpectedString);
         actualBool.Should().Be(ExpectedBool);
         actualDate.Should().BeCloseTo(ExpectedDateTime);
         actualObject.Should().Be(ExpectedObject);
      }

      [TestMethod]
      public void When_I_override_a_global_method_manually()
      {
         // Arrange.
         const string ExpectedText = "Echo from C#";
         string actualText = null;

         scriptEngine.AddCode(
            @"Public Sub Echo(sText)
                 ' Does nothing
              End Sub");

         FunctionOverrider overrider = new FunctionOverrider();

         Action<object[]> echoAction = (args) => actualText = args[0].ToString();
         overrider.WhenCalled(echoAction);

         // Act.
         scriptEngine.AddObject("Echo", overrider);
         script.Echo(ExpectedText);

         // Assert.
         actualText.Should().BeEquivalentTo(ExpectedText);
      }

      [TestMethod]
      public void When_I_override_a_global_method()
      {
         // Arrange.
         const string ExpectedText = "Echo from C#";
         string actualText = null;

         scriptEngine.AddCode(
            @"Public Sub Echo(sText)
                 ' Does nothing
              End Sub");

         Action<object[]> echoAction = (args) => actualText = args[0].ToString();

         // Act.
         scriptEngine.OverrideGlobalFunction("Echo", echoAction);
         script.Echo(ExpectedText);

         // Assert.
         actualText.Should().BeEquivalentTo(ExpectedText);
      }

      [TestMethod]
      public void When_I_override_a_global_function_manually()
      {
         // Arrange.
         scriptEngine.AddCode(
            @"Public Function Add(a, b)
                 Err.Raise 1
              End Function");

         FunctionOverrider overrider = new FunctionOverrider();

         Func<object[], object> addAction = (args) => (int)args[0] + (int)args[1];
         overrider.WhenCalled(addAction);

         // Act.
         scriptEngine.AddObject("Add", overrider);
         int result = script.Add(5, 7);

         // Assert.
         result.Should().Be(12);
      }

      [TestMethod]
      public void When_I_override_a_global_function()
      {
         // Arrange.
         scriptEngine.AddCode(
            @"Public Function Add(a, b)
                 Err.Raise 1
              End Function");

         Func<object[], object> addAction = (args) => (int)args[0] + (int)args[1];

         // Act.
         scriptEngine.OverrideGlobalFunction("Add", addAction);
         int result = script.Add(5, 7);

         // Assert.
         result.Should().Be(12);
      }
   }
}