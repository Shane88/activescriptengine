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
      private string scriptLanguageProgId = VBScript.ProgId;
      private bool shouldStart;

      /// <summary>
      /// Configures the ProgId of the language to use for the engine.
      /// If not specified VBScript is used by default.
      /// </summary>
      /// <param name="progId">The progId of the language to use for the engine.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder UseLanguage(string progId)
      {
         scriptLanguageProgId = progId;
         return this;
      }

      /// <summary>
      /// Allows custom configuration of the ActiveScriptEngine being built.
      /// </summary>
      /// <param name="configuration">The configuration action that should occur after the engine is built.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder Configure(Action<ActiveScriptEngine> configuration)
      {
         configurationActions.Add(configuration);
         return this;
      }

      /// <summary>
      /// Adds any code files that match the specified search pattern. The search pattern can be a relative path and can
      /// navigate up directories with "..\". A * wild card character is also supported for the file name.
      /// A searchPattern of "..\..\Code\*.vbs would navigate up two directories, then into the Code directory and then
      /// add all the files with an extension of .vbs.
      /// </summary>
      /// <param name="searchPattern">The search pattern of the files to add.</param>
      /// <param name="namespaceName">The namespace that the code files should be added to. If not specified they will be added to the root namespace.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder AddCodeFiles(string searchPattern, string namespaceName = null)
      {
         IEnumerable<string> filePaths = FindFiles(searchPattern);

         foreach (string filePath in filePaths)
         {
            AddCode(File.ReadAllText(filePath), Path.GetFileName(filePath), namespaceName);
         }

         return this;
      }

      /// <summary>
      /// Adds the specified code to the script engine.
      /// </summary>
      /// <param name="code">The code to add.</param>
      /// <param name="scriptName">The name of the script. This could be a file name or any name to identify the script.
      /// This is useful for troubleshooting purposes because the script name is given to your during errors.</param>
      /// <param name="namespaceName">The namespace to add the code too.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
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

      /// <summary>
      /// Adds the specified object to the script engine with the specified name.
      /// </summary>
      /// <param name="name"></param>
      /// <param name="value"></param>
      /// <returns></returns>
      public ActiveScriptEngineBuilder AddObject<T>(string name, T value) where T : class
      {
         return AddObject(name, () => value);
      }

      public ActiveScriptEngineBuilder AddObject<T>(string name) where T : class, new()
      {
         return AddObject(name, () => new T());
      }

      /// <summary>
      /// Creates an instance of type T and adds it to the engine with the specified name.
      /// Allows additional configuration to be performed on the created instance by supplying
      /// an action to perform on the instance.
      /// </summary>
      /// <typeparam name="T">The type of the object to create. Must be a class with a default constructor.</typeparam>
      /// <param name="name">The name of the object once it is added to the ActiveScriptEngine.</param>
      /// <param name="configurationAction">An action which will receive the created instance of type T
      /// which can perform addition setup of the instance.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
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

      public ActiveScriptEngineBuilder AddObject<T>(string name, Func<T> factory) where T : class
      {
         if (objectFactories.ContainsKey(name))
         {
            Func<object> existingFactory = objectFactories[name];

            Type funcFactoryType = existingFactory.GetType().GetGenericArguments()[0];

            string exceptionMessage = string.Format(
               "Cannot add object with name {0} and type {1} to the builder because an object has already been added with that name with type of {2}",
               name,
               typeof(T).Name,
               funcFactoryType.Name);

            throw new ArgumentException(exceptionMessage, "name");
         }

         // where T : class is required for the implicit cast of Func<T> to Func<object>
         // when adding the func to the dictionary.
         objectFactories.Add(name, factory);
         return this;
      }

      public ActiveScriptEngine Build()
      {
         var engine = new ActiveScriptEngine(scriptLanguageProgId);

         RunConfigurations(engine);

         AddObjectsToEngine(engine);

         AddCodeToEngine(engine);

         if (shouldStart)
         {
            engine.Start();
         }

         return engine;
      }

      private void AddCodeToEngine(ActiveScriptEngine engine)
      {
         foreach (CodeBlock codeBlock in codeBlocks)
         {
            try
            {
               engine.AddCode(codeBlock.CodeText, codeBlock.Namespace, codeBlock.ScriptName);
            }
            catch (Exception ex)
            {
               throw new InvalidOperationException(
                  "An exception occurred when trying to add code with name " + (codeBlock.ScriptName ?? "Untitled") +
                  ". Check the InnerException for more details",
                  ex);
            }
         }
      }

      private void AddObjectsToEngine(ActiveScriptEngine engine)
      {
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
      }

      private void RunConfigurations(ActiveScriptEngine engine)
      {
         foreach (Action<ActiveScriptEngine> action in configurationActions)
         {
            try
            {
               action(engine);
            }
            catch (Exception ex)
            {
               throw new InvalidOperationException(
                  "An exception occurred when trying to Configure the script engine. Check the InnerException for more details",
                  ex);
            }
         }
      }

      public ActiveScriptEngineBuilder LogErrorsToTrace()
      {
         return LogErrorsTo(errorText => Trace.WriteLine(errorText));
      }

      public ActiveScriptEngineBuilder LogErrorsTo(Action<string> logAction)
      {
         return Configure(
            engine =>
               engine.ScriptErrorOccurred += (sender, error) =>
                  logAction(FormatErrorInfo(error)));
      }

      private static string FormatErrorInfo(ScriptErrorInfo error)
      {
         return string.Format(
            "Error in {0} on Ln {1} Col {2}, {3}{4}Error Code:{5}, Text:{6}",
            error.ScriptName ?? "Unnamed Script",
            error.LineNumber,
            error.ColumnNumber,
            error.Description,
            Environment.NewLine,
            error.ErrorNumber,
            error.LineText);
      }

      public ActiveScriptEngineBuilder StartEngineOnBuild()
      {
         shouldStart = true;
         return this;
      }

      private static IEnumerable<string> FindFiles(string searchPattern)
      {
         return FindFiles(searchPattern, Environment.CurrentDirectory);
      }

      private static IEnumerable<string> FindFiles(string searchPattern, string rootDirectoryPath)
      {
         // Get directory and file parts of complete relative pattern.
         string filePattern = Path.GetFileName(searchPattern);
         string relativePath = Path.GetDirectoryName(searchPattern);

         string absolutePath = Path.IsPathRooted(rootDirectoryPath) ?
                                  relativePath :
                                  Path.GetFullPath(Path.Combine(rootDirectoryPath, relativePath));

         // Search files matching the pattern.
         return Directory.EnumerateFiles(absolutePath, filePattern, SearchOption.TopDirectoryOnly);
      }

      private class CodeBlock
      {
         public string CodeText { get; set; }

         public string Namespace { get; set; }

         public string ScriptName { get; set; }
      }
   }
}
