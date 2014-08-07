namespace ActiveXScriptLibUnitTests
{
   using System;
   using System.Collections.Generic;
   using System.Runtime.InteropServices;
   using ActiveXScriptLib;
   using ActiveXScriptLib.Extensions;
   using FluentAssertions;
   using Microsoft.CSharp.RuntimeBinder;
   using Microsoft.VisualStudio.TestTools.UnitTesting;

   [TestClass]
   public class VBScriptComProxyTests : VBScriptTestBase
   {
      public class MyComInvisibleClass
      {
         public string Name { get; set; }
      }

      [ComVisible(true)]
      public class MyComVisibleClass
      {
         public string Name { get; set; }
      }

      [TestMethod]
      public void When_a_non_com_visible_dot_net_object_is_added_and_used_from_VBScript()
      {
         // Arrange.
         var mtc = new MyComInvisibleClass()
         {
            Name = "Hello World"
         };

         scriptEngine.AddObject("mtc", mtc);

         scriptEngine.Start();

         // Act.
         string name = scriptEngine.Evaluate<string>("mtc.Name");

         // Assert.
         name.Should().Be("Hello World");
      }

      [TestMethod]
      public void When_a_non_com_visible_dot_net_object_is_added_and_used_from_csharp()
      {
         // Arrange.
         var mtc = new MyComInvisibleClass()
         {
            Name = "Hello World"
         };

         scriptEngine.AddObject("mtc", mtc);

         scriptEngine.Start();

         string name = null;

         // Act.
         Action action = () => name = script.mtc.Name;

         // Assert.
         action.ShouldThrow<RuntimeBinderException>();
      }

      [TestMethod]
      public void When_a_com_visible_dot_net_object_is_added_and_used_from_vbscript()
      {
         // Arrange.
         var mtc = new MyComVisibleClass()
         {
            Name = "Hello World"
         };

         scriptEngine.AddObject("mtc", mtc);

         scriptEngine.Start();

         // Act.
         string name = scriptEngine.Evaluate<string>("mtc.Name");

         // Assert.
         name.Should().Be("Hello World");
      }

      [TestMethod]
      public void When_a_com_visible_dot_net_object_is_added_and_used_from_csharp()
      {
         // Arrange.
         var mtc = new MyComVisibleClass()
         {
            Name = "Hello World"
         };

         scriptEngine.AddObject("mtc", mtc);

         scriptEngine.Start();

         // Act.
         string name = script.mtc.Name;

         // Assert.
         name.Should().Be("Hello World");
      }

      [TestMethod]
      public void When_a_real_com_object_is_added_and_used_from_vbscript()
      {
         // Arrange.
         dynamic comDictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
         comDictionary.Add(1, "One");

         scriptEngine.AddObject("items", comDictionary);

         scriptEngine.Start();

         // Act.
         string value = scriptEngine.Evaluate<string>("items(1)");

         // Assert.
         value.Should().Be("One");
      }

      [TestMethod]
      public void When_a_real_com_object_is_added_and_used_from_csharp()
      {
         // Arrange.
         dynamic comDictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
         comDictionary.Add(1, "One");

         scriptEngine.AddObject("items", comDictionary);

         scriptEngine.Start();

         // Act.
         string value = script.items[1];

         // Assert.
         value.Should().Be("One");
      }

      [TestMethod]
      public void When_a_COMProxy_is_added_and_used_from_vbscript()
      {
         // Arrange.
         var mtc = new MyComInvisibleClass()
         {
            Name = "Hello World"
         };

         var comProxy = new ComProxy(mtc);

         scriptEngine.AddObject("mtc", comProxy);

         scriptEngine.Start();

         // Act.
         string name = scriptEngine.Evaluate<string>("mtc.Name");

         // Assert.
         name.Should().Be("Hello World");
      }

      [TestMethod]
      public void When_a_COMProxy_is_added_and_used_from_csharp()
      {
         // Arrange.
         var mtc = new MyComInvisibleClass()
         {
            Name = "Hello World"
         };

         var comProxy = new ComProxy(mtc);

         scriptEngine.AddObject("mtc", comProxy);

         scriptEngine.Start();

         string name = null;

         // Act.
         Action action = () => name = script.mtc.Name;

         // Assert.
         action.ShouldThrow<RuntimeBinderException>();
      }

      [TestMethod]
      public void When_a_GenericDictionary_is_added()
      {
         // Arrange.
         var dictionary = new Dictionary<int, string>();

         scriptEngine.Start();

         scriptEngine.AddObject("values", dictionary);

         dictionary.Add(1, "One");

         // Act.
         scriptEngine.AddCode("values.Add 2, \"Two\"");

         int count = scriptEngine.Evaluate<int>("values.Count");

         // Assert.
         count.Should().Be(2);
      }

      [TestMethod]
      public void When_a_GenericDictionary_is_passed_as_a_parameter()
      {
         // Arrange.
         var dictionary = new Dictionary<int, string>();

         scriptEngine.Start();
         
         dictionary.Add(1, "One");
         

         scriptEngine.AddCode(@"
Public Function PrintValues(values)

   values.Add 2, ""Two""

   Dim oEnum
   Set oEnum = values

   WScript.Echo TypeName(oEnum)

   Dim val
   For Each val In oEnum
      WScript.Echo val
   Next

   PrintValues = values.Count
End Function");

         // Act.
         int count = script.PrintValues(new ComProxy(dictionary));

         // Assert.
         count.Should().Be(2);
      }
   }
}