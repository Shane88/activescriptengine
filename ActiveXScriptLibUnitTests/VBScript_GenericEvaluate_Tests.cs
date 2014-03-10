namespace ActiveXScriptLibUnitTests
{
   using System;
   using ActiveXScriptLib;
   using ActiveXScriptLib.Extensions;
   using FluentAssertions;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class VBScript_GenericEvaluate_Tests : VBScriptTestBase
   {
      [TestInitialize]
      public void TestSetup()
      {
         scriptEngine.AddCode(
            @"Dim value

              Public Sub SetValue(val)
                 If IsObject(val) Then
                    Set value = val
                 Else
                    value = val
                 End If
              End Sub

              Public Function GetValue()
                 If IsObject(value) Then
                    Set Getvalue = value
                 Else
                    GetValue = value
                 End If
              End Function");

         scriptEngine.Start();
      }

      [TestMethod]
      public void When_GenericEvaluate_is_called_with_a_null_engine()
      {
         // Arrange.
         ActiveScriptEngine engine = null;

         // Act.
         Action action = () => engine.Evaluate<object>("GetValue");

         // Assert.
         action.ShouldThrow<ArgumentNullException>();
      }

      [TestMethod]
      public void When_GenericEvaluate_is_called_with_various_types()
      {
         EvaluateTest<string, string>("123", "123");
         EvaluateTest<string, int>("123", 123);
         EvaluateTest<short, int>((short)2, 2);
         EvaluateTest<string, double>("53435", 53435.0);
         EvaluateTest<int, decimal>(53435, 53435.0m);
         EvaluateTest<short, uint>(123, 123u);
         EvaluateTest<float, int>(123.3f, 123);
      }

      private void EvaluateTest<TInput, TOutput>(TInput inputValue, TOutput expectedOutput)
      {
         // Arrange.
         script.SetValue(inputValue);

         // Act.
         var outputValue = scriptEngine.Evaluate<TOutput>("GetValue");

         // Assert.
         outputValue.Should().Be(expectedOutput);
      }
   }
}