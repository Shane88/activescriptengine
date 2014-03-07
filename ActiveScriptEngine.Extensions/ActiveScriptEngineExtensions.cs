namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Globalization;

   /// <summary>
   /// Provides useful extension methods when using the ActiveScriptEngine.
   /// </summary>
   public static class ActiveScriptEngineExtensions
   {
      /// <summary>
      /// Creates a new instance of the specified class from within the script engine and returns it.
      /// </summary>
      /// <param name="engine">The ActiveScriptEngine to create the instance from.</param>
      /// <param name="className">The name of the class to create an instance of.</param>
      /// <returns>The created instance.</returns>
      public static object NewInstanceOf(this ActiveScriptEngine engine, string className)
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }

         if (engine.ProgId.Equals(VBScript.ProgId, StringComparison.OrdinalIgnoreCase))
         {
            return engine.Evaluate("New " + className);
         }

         return engine.Evaluate("new " + className + "()");
      }

      /// <summary>
      /// Provides a generic overload for engine.Evaluate.
      /// This method will invoke Convert.ChangeType on the returned value to try convert it to type T.
      /// This provides a more easier way to get the value of a type you want.
      /// </summary>
      /// <typeparam name="T">The Type to convert the returned value of Evaluate to.</typeparam>
      /// <param name="engine">The ActiveScriptEngine to invoke Evaluate on.</param>
      /// <param name="code">The code to evaluate inside engine.</param>
      /// <returns>The value returned after the evaluation as type T.</returns>
      public static T Evaluate<T>(this ActiveScriptEngine engine, string code)
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }

         return (T)Convert.ChangeType(engine.Evaluate(code), typeof(T), CultureInfo.InvariantCulture);
      }
   }
}