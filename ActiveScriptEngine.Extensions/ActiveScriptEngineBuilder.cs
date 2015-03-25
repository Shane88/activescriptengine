namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Collections.Generic;
   using System.Diagnostics;
   using System.IO;

   public class ActiveScriptEngineBuilder
   {
      private readonly List<CodeBlock> codeBlocks = new List<CodeBlock>();
      private readonly List<Action<ActiveScriptEngine>> configurationActions = new List<Action<ActiveScriptEngine>>();
      private readonly Dictionary<string, Func<object>> objectFactories = new Dictionary<string, Func<object>>();
      private CreateObjectFactory createObjectFactory;
      private bool shouldStart;

      public ActiveScriptEngineBuilder Configure(Action<ActiveScriptEngine> configuration)
      {
         configurationActions.Add(configuration);
         return this;
      }

      public ActiveScriptEngineBuilder LogErrorsToTrace()
      {
         Configure(
            engine =>
            {
               engine.ScriptErrorOccurred += (sender, error) =>
               {
                  Trace.WriteLine(
                     string.Format(
                        "Error in {0} on Ln {1} Col {2}, {3}{4}Error Code:{5}, Text:{6}",
                         error.ScriptName ?? "Unnamed Script",
                         error.LineNumber,
                         error.ColumnNumber,
                         error.Description,
                         Environment.NewLine,
                         error.ErrorNumber,
                         error.LineText));
               };
            });

         return this;
      }

      public ActiveScriptEngineBuilder AddCodeFile(string filePath, string namespaceName = null)
      {
         return AddCode(File.ReadAllText(filePath), Path.GetFileName(filePath), null);
      }

      public ActiveScriptEngineBuilder AddCode(string code, string scriptName = null, string namespaceName = null)
      {
         CodeBlock codeBlock = new CodeBlock()
         {
            CodeText = code,
            ScriptName = scriptName,
            Namespace = namespaceName
         };

         codeBlocks.Add(codeBlock);

         return this;
      }

      public ActiveScriptEngineBuilder AddObject(string name, object value)
      {
         return AddObject(name, () => value);
      }

      public ActiveScriptEngineBuilder AddObject<T>(string name) where T : class, new()
      {
         return AddObject(name, () => new T());
      }

      public ActiveScriptEngineBuilder AddObject<T>(string name, Func<T> factory) where T : class
      {
         // Cast is required so we can call the correct AddObject(string, Func<object>) overload
         // which adds to a dictionary of the same types.
         // where T : class is required for the cast to work.
         return AddObject(name, (Func<object>)factory);
      }

      public ActiveScriptEngineBuilder AddObject(string name, Func<object> factory)
      {
         objectFactories.Add(name, factory);
         return this;
      }

      public ActiveScriptEngineBuilder AddObject<T>(string name, Action<T> configurationAction) where T : class, new()
      {
         Func<object> factory = () =>
         {
            T instance = new T();
            configurationAction(instance);
            return instance;
         };

         objectFactories.Add(name, factory);
         return this;
      }

      public ActiveScriptEngine Build()
      {
         var engine = new ActiveScriptEngine(VBScript.ProgId);

         foreach (var action in configurationActions)
         {
            action(engine);
         }

         foreach (string objectName in objectFactories.Keys)
         {
            Func<object> objectFactory = objectFactories[objectName];

            try
            {
               object objectToAdd = objectFactory();
               engine.AddObject(objectName, objectToAdd);
            }
            catch (Exception ex)
            {
               throw new InvalidOperationException(
                  "An exception occurred when trying to add object with name " + objectName + ". Check the InnerException for more details",
                  ex);
            }
         }

         foreach (CodeBlock codeBlock in codeBlocks)
         {
            engine.AddCode(codeBlock.CodeText, codeBlock.Namespace, codeBlock.ScriptName);
         }

         if (shouldStart)
         {
            engine.Start();
         }

         return engine;
      }

      public ActiveScriptEngineBuilder AddCreateObjectHook(Action<CreateObjectFactory> factory)
      {
         if (createObjectFactory == null)
         {
            createObjectFactory = new CreateObjectFactory();
            AddObject("CreateObject", createObjectFactory);
         }

         factory(createObjectFactory);

         return this;
      }

      public ActiveScriptEngineBuilder StartEngineOnBuild()
      {
         shouldStart = true;
         return this;
      }

      private class CodeBlock
      {
         public string CodeText { get; set; }

         public string Namespace { get; set; }

         public string ScriptName { get; set; }
      }
   }
}
