namespace ActiveXScriptLib
{
   using System;
   using System.Collections.Generic;
   using System.Diagnostics;
   using System.Runtime.InteropServices;
   using Interop.ActiveXScript;
   using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

   public class ActiveScriptEngine : IDisposable
   {
      private IActiveScript activeScript;
      private Dictionary<string, object> hostObjects;

      private ActiveScriptParse parser;
      private ActiveScriptSite scriptSite;

      private Dictionary<ulong, ScriptInfo> scripts;
      private string scriptToParse;
      private string progId;

      public ActiveScriptEngine(string progId)
      {
         if (progId == null)
         {
            throw new ArgumentNullException("progId");
         }

         if (string.IsNullOrEmpty(progId))
         {
            throw new ArgumentException("progID must not be empty");
         }

         Type type = Type.GetTypeFromProgID(progId, true);

         Debug.Assert(type != null, "Type should be valid at this point as the above line will throw if the progId is invalid");

         object activeScriptInstance = Activator.CreateInstance(type);

         this.Initialise(activeScriptInstance);
         this.progId = progId;
      }

      public ActiveScriptEngine(object activeScriptInstance)
      {
         if (activeScriptInstance == null)
         {
            throw new ArgumentNullException("activeScriptInstance");
         }

         this.Initialise(activeScriptInstance);

         // TODO: It would be nice it we could find the progId here.
      }

      ~ActiveScriptEngine()
      {
         Dispose(false);
      }

      public event ScriptErrorOccurredEventHandler ScriptErrorOccurred;

      public string ProgId
      {
         get
         {
            return this.progId;
         }
      }

      public ScriptErrorInfo LastError { get; internal set; }

      /// <summary>
      /// Gets a value indicating whether the script engine is running.
      /// </summary>
      public bool IsRunning
      {
         get
         {
            return activeScript.GetScriptState() == ScriptState.Started;
         }
      }

      internal IActiveScript ActiveScript
      {
         get
         {
            return this.activeScript;
         }
      }

      internal Dictionary<string, object> HostObjects
      {
         get
         {
            return this.hostObjects;
         }
      }

      /// <summary>
      /// Adds the specified code to the scripting engine.
      /// </summary>
      /// <param name="code">The code to be added.</param>
      /// <exception cref="ArgumentNullException">If code is null.</exception>
      /// <exception cref="ArgumentException">If code is blank.</exception>
      public void AddCode(string code)
      {
         AddCode(code, null);
      }

      /// <summary>
      /// Adds the specified code to the scripting engine.
      /// The code that is added will be available only under the namespace specified
      /// instead of in the global scope of the script.
      /// </summary>
      /// <param name="code">The code to be added.</param>
      /// <param name="namespaceName">The name of the namespace to add the code to.</param>
      /// <exception cref="ArgumentNullException">If code is null.</exception>
      /// <exception cref="ArgumentException">If code is blank.</exception>
      public void AddCode(string code, string namespaceName)
      {
         AddCode(code, namespaceName, null);
      }

      /// <summary>
      /// Adds the specified code to the scripting engine under the specified namespace.
      /// The code that is added will be available only under the namespace specified
      /// instead of in the global scope of the script. The script name will be used
      /// to provide more useful error information if an error occurs during the execution
      /// of the code that was added.
      /// </summary>
      /// <param name="code">The code to be added.</param>
      /// <param name="namespaceName">The name of the namespace to add the code to.</param>
      /// <param name="scriptName">The script name that the code came from.</param>
      /// <exception cref="ArgumentNullException">If code is null.</exception>
      /// <exception cref="ArgumentException">If code is blank.</exception>
      public void AddCode(string code, string namespaceName, string scriptName)
      {
         if (code == null)
         {
            throw new ArgumentNullException("code");
         }

         if (string.IsNullOrEmpty(code))
         {
            throw new ArgumentException("code parameter must contain code", "code");
         }

         if (namespaceName != null)
         {
            activeScript.AddNamedItem(namespaceName, ScriptItemFlags.CodeOnly | ScriptItemFlags.IsVisible);
         }

         try
         {
            /*
             * In the event that the passed in script is not valid syntax
             * an error will be thrown by the script engine. This will be
             * handled by the OnScriptError event before the exception
             * is thrown here so we need to set this variable to use
             * in the OnScriptError block to figure out the script name
             * since it won't have been added to the script list.
             */
            scriptToParse = scriptName;

            EXCEPINFO exceptionInfo = new EXCEPINFO();

            ulong cookie = (ulong)scripts.Count;

            parser.ParseScriptText(
                code: code,
                itemName: namespaceName,
                context: null,
                delimiter: null,
                sourceContext: cookie,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsVisible,
                pVarResult: IntPtr.Zero,
                excepInfo: out exceptionInfo);

            ScriptInfo si = new ScriptInfo()
            {
               Code = code,
               ScriptName = scriptName
            };

            scripts.Add(cookie, si);
         }
         finally
         {
            scriptToParse = null;
         }
      }

      /// <summary>
      /// Adds the specified object into the script context with the specified name.
      /// The methods on the object can be called without the name as if they were
      /// native functions available in the script.
      /// </summary>
      /// <param name="name">The name the object can optionally be referenced as.</param>
      /// <param name="value">The object to be added.</param>
      /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
      public void AddGlobalMemberObject(string name, object value)
      {
         if (hostObjects.ContainsKey(name))
         {
            throw new ArgumentException("object with that name already exists", "name");
         }

         hostObjects.Add(name, value);
         activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible | ScriptItemFlags.GlobalMembers);
      }

      /// <summary>
      /// Adds the specified object into the script context with the specified name.
      /// </summary>
      /// <param name="name">The name the object will be referenced as.</param>
      /// <param name="value">The object to be added.</param>
      /// <exception cref="ArgumentException">If the name has been used previously to add an object.</exception>
      public void AddObject(string name, object value)
      {
         if (hostObjects.ContainsKey(name))
         {
            throw new ArgumentException("object with that name already exists", "name");
         }

         if (Marshal.IsComObject(value) ||
            IsComVisibleDotNetObject(value) ||
            value is ComProxy)
         {
            hostObjects.Add(name, value);
         }
         else
         {
            ComProxy proxy = new ComProxy(value);
            hostObjects.Add(name, proxy);
         }

         activeScript.AddNamedItem(name, ScriptItemFlags.IsSource | ScriptItemFlags.IsVisible);
      }

      private static bool IsComVisibleDotNetObject(object value)
      {
         bool comVisible = false;

         Type type = value.GetType();

         var comVisibleAttributes = 
            type.GetCustomAttributes(typeof(ComVisibleAttribute), false) as ComVisibleAttribute[];

         if (comVisibleAttributes != null)
         {
            foreach (var comAttribute in comVisibleAttributes)
            {
               comVisible = comAttribute.Value;
               break;
            }
         }

         return comVisible;
      }

      /// <summary>
      /// Gets the IDispatch handle for the root namespace.
      /// </summary>
      /// <returns>An IDispatch handle for the root namespace.</returns>
      public object GetScriptHandle()
      {
         return GetScriptHandle(null);
      }

      /// <summary>
      /// Gets the IDispatch handle for the specified namespace.
      /// </summary>
      /// <param name="namespaceName">The namespace of the IDispatch handle to get, or null
      /// to get the root namespace.</param>
      /// <returns>An IDispatch handle.</returns>
      public object GetScriptHandle(string namespaceName)
      {
         return activeScript.GetScriptDispatch(namespaceName);
      }

      /// <summary>
      /// Gets the script host object of the specified name.
      /// </summary>
      /// <param name="name">The name of the object.</param>
      /// <returns>The script host object.</returns>
      /// <exception cref="ArgumentException">If the specified host object did not exist.</exception>
      public object GetScriptObject(string name)
      {
         if (!ScriptHasHostObject(name))
         {
            throw new ArgumentException("object with that name does not exist", "name");
         }

         return hostObjects[name];
      }

      /// <summary>
      /// Evaluates the specified code inside the context of the script and returns
      /// the result, or null if no result was returned.
      /// </summary>
      /// <param name="code">The code to evaluate.</param>
      /// <returns>The result of the evaluation.</returns>
      /// <exception cref="ArgumentNullException">If code is null.</exception>
      /// <exception cref="ArgumentException">If code is blank.</exception>
      public object Evaluate(string code)
      {
         if (code == null)
         {
            throw new ArgumentNullException("code");
         }

         if (string.IsNullOrEmpty(code))
         {
            throw new ArgumentException("code parameter must contain code", "code");
         }

         EXCEPINFO exceptionInfo = new EXCEPINFO();

         ulong cookie = (ulong)scripts.Count;

         IntPtr pResult = Marshal.AllocCoTaskMem(1024);

         try
         {
            parser.ParseScriptText(
                code: code,
                itemName: null,
                context: null,
                delimiter: null,
                sourceContext: cookie,
                startingLineNumber: 1u,
                flags: ScriptTextFlags.IsExpression,
                pVarResult: pResult,
                excepInfo: out exceptionInfo);

            if (pResult != IntPtr.Zero)
            {
               return Marshal.GetObjectForNativeVariant(pResult);
            }
         }
         finally
         {
            Marshal.FreeCoTaskMem(pResult);
         }

         return null;
      }

      /// <summary>
      /// Puts the scripting engine into the Started state.
      /// At this point code will be executed in the order they were added to the script engine.
      /// </summary>
      public void Start()
      {
         activeScript.SetScriptState(ScriptState.Started);
         activeScript.SetScriptState(ScriptState.Connected);
      }

      /// <summary>
      /// Stops and disposes this ActiveScriptEngine.
      /// </summary>
      public void Dispose()
      {
         Dispose(true);
         GC.SuppressFinalize(this);
      }

      /// <summary>
      /// Returns if the script contains a host object of the specified name.
      /// </summary>
      /// <param name="name">The name of the object.</param>
      /// <returns>True if the script contains a host object of the specified name, otherwise false.</returns>
      public bool ScriptHasHostObject(string name)
      {
         return hostObjects.ContainsKey(name);
      }

      internal void OnScriptError(IActiveScriptError error)
      {
         this.LastError = new ScriptErrorInfo(error, this.scripts);

         if (this.LastError.ScriptName == null && scriptToParse != null)
         {
            this.LastError.ScriptName = scriptToParse;
         }

         var errorOccurred = this.ScriptErrorOccurred;

         if (errorOccurred != null)
         {
            errorOccurred(this, this.LastError);
         }
      }

      protected virtual void Dispose(bool managedDispose)
      {
         if (activeScript != null)
         {
            try
            {
               activeScript.Close();
            }
            finally
            {
               activeScript = null;
            }
         }

         if (parser != null)
         {
            parser = null;
         }

         if (scriptSite != null)
         {
            scriptSite = null;
         }
      }

      private void Initialise(object activeScriptObject)
      {
         IActiveScript activeScriptInstance = activeScriptObject as IActiveScript;

         if (activeScriptInstance == null)
         {
            throw new InvalidActiveScriptClassException();
         }

         this.parser = ActiveScriptParse.MakeActiveScriptParse(activeScriptInstance);

         this.parser.InitNew();

         this.scriptSite = new ActiveScriptSite(this);
         activeScriptInstance.SetScriptSite(this.scriptSite);

         this.hostObjects = new Dictionary<string, object>();
         this.scripts = new Dictionary<ulong, ScriptInfo>();

         this.activeScript = activeScriptInstance;
      }
   }
}