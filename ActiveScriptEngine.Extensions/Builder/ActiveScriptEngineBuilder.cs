namespace ActiveXScriptLib.Extensions.Builder
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
      private readonly List<object> actionsToRunOnStart = new List<object>();

      private string scriptLanguageProgId = VBScript.ProgId;
      private bool shouldStart;

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
      /// Adds any code files that match the specified search pattern. The search pattern can be a relative path and can
      /// navigate up directories with "..\". A * wild card character is also supported for the file name.
      /// A searchPattern of "..\..\Code\*.vbs would navigate up two directories, then into the Code directory and then
      /// add all the files with an extension of .vbs.
      /// </summary>
      /// <param name="searchPattern">The search pattern of the files to add.</param>
      /// <param name="namespaceName">The namespace that the code files should be added to. If not specified they will be added to the root namespace.</param>
      /// <param name="rootDirectory">The root directory of the files to search for. If not specified Environment.CurrentDirectory is used.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder AddCodeFiles(string searchPattern, string namespaceName = null, string rootDirectory = null)
      {
         IEnumerable<string> filePaths = rootDirectory == null ?
            FindFiles(searchPattern) :
            FindFiles(searchPattern, rootDirectory);

         foreach (string filePath in filePaths)
         {
            AddCode(File.ReadAllText(filePath), Path.GetFileName(filePath), namespaceName);
         }

         return this;
      }

      /// <summary>
      /// Adds the specified object into the script context with the specified name.
      /// </summary>
      /// <typeparam name="T">The type of the object to add.</typeparam>
      /// <param name="name">The name the object will be referenced as.</param>
      /// <param name="value">The object to be added.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
      public ActiveScriptEngineBuilder AddObject<T>(string name, T value) where T : class
      {
         return AddObject(name, () => value);
      }

      /// <summary>
      /// Adds a new object of type T into the script context with the specified name.
      /// </summary>
      /// <typeparam name="T">The type of the object to create and add.</typeparam>
      /// <param name="name">The name the object will be referenced as.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
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

      /// <summary>
      /// Adds the object returned by objectFactory to the engine with the specified name.
      /// </summary>
      /// <typeparam name="T">The type of the object to add.</typeparam>
      /// <param name="name">The name of the object once it is added to the ActiveScriptEngine.</param>
      /// <param name="objectFactory">The factory which produces an instance of type T.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder AddObject<T>(string name, Func<T> objectFactory) where T : class
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
         objectFactories.Add(name, objectFactory);
         return this;
      }

      /// <summary>
      /// Builds the ActiveScriptEngine using the configuration configured with this builder.
      /// </summary>
      /// <returns>A newly built ActiveScriptEngine instance.</returns>
      public ActiveScriptEngine Build()
      {
         var engine = new ActiveScriptEngine(scriptLanguageProgId);

         RunConfigurations(engine);

         AddObjectsToEngine(engine);

         AddCodeToEngine(engine);

         if (shouldStart)
         {
            engine.Start();
            RunActionsOnStart(engine);
         }

         return engine;
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
      /// Configures a event handler for the ScriptErrorOccurred and logs the output to Trace.
      /// </summary>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder LogErrorsToTrace()
      {
         return LogErrorsTo(errorText => Trace.WriteLine(errorText));
      }

      /// <summary>
      /// Configures an event handler for the ScriptErrorOccurred event which formats an error message and calls the supplied logAction.
      /// </summary>
      /// <param name="logAction">The action to delegate the text of scripts errors to.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder LogErrorsTo(Action<string> logAction)
      {
         return Configure(
            engine =>
               engine.ScriptErrorOccurred += (sender, error) =>
                  logAction(FormatErrorInfo(error)));
      }

      /// <summary>
      /// Configures an action to be called after the engine has been built.
      /// This action will receive a dynamic which is an instance to the root script handle.
      /// Note: Calling this method will automatically configure the engine to be started when it is built.
      /// This is the same as calling StartEngineOnBuild explicitly.
      /// </summary>
      /// <param name="onStartAction">The action to run once the engine is built.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder OnStart(Action<dynamic> onStartAction)
      {
         this.StartEngineOnBuild();
         actionsToRunOnStart.Add(onStartAction);
         return this;
      }

      /// <summary>
      /// Configures code to be run after the engine has been built.
      /// Note: Calling this method will automatically configure the engine to be started when it is built.
      /// This is the same as calling StartEngineOnBuild explicitly.
      /// </summary>
      /// <param name="code">The code to run.</param>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder RunCodeOnStart(string code)
      {
         this.StartEngineOnBuild();
         actionsToRunOnStart.Add(code);
         return this;
      }

      /// <summary>
      /// Configures the engine to be started upon being successfully built.
      /// </summary>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder StartEngineOnBuild()
      {
         shouldStart = true;
         return this;
      }

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
      /// Configures the engine to use VBScript as the programming language.
      /// </summary>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder UseVBScript()
      {
         scriptLanguageProgId = VBScript.ProgId;
         return this;
      }

      /// <summary>
      /// Configures the engine to use JScript as the programming language.
      /// </summary>
      /// <returns>This ActiveScriptEngineBuilder to allow for fluent method calls.</returns>
      public ActiveScriptEngineBuilder UseJScript()
      {
         scriptLanguageProgId = JScript.ProgId;
         return this;
      }

      /// <summary>
      /// Adds all code which has been added to the builder to the engine.
      /// </summary>
      /// <param name="engine">The engine to add the code to.</param>
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

      /// <summary>
      /// Adds all objects which have been added to the builder to the engine.
      /// </summary>
      /// <param name="engine">The engine to add the objects to.</param>
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

      private void RunActionsOnStart(ActiveScriptEngine engine)
      {
         foreach (object action in actionsToRunOnStart)
         {
            string code = action as string;

            if (code != null)
            {
               HandleAddCodeOnStart(engine, code);
               continue;
            }

            Action<dynamic> actionOnStart = action as Action<dynamic>;

            if (actionOnStart != null)
            {
               HandleActionOnStart(engine, actionOnStart);
               continue;
            }
         }
      }

      private void HandleActionOnStart(ActiveScriptEngine engine, Action<dynamic> actionOnStart)
      {
         dynamic scriptHandle = engine.GetScriptHandle();
         actionOnStart(scriptHandle);
      }

      private void HandleAddCodeOnStart(ActiveScriptEngine engine, string code)
      {
         engine.AddCode(code);
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

      private class CodeBlock
      {
         public string CodeText { get; set; }

         public string Namespace { get; set; }

         public string ScriptName { get; set; }
      }
   }
}
