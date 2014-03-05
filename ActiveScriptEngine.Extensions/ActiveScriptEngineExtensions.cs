namespace ActiveXScriptLib.Extensions
{
   using System;

   public static class ActiveScriptEngineExtensions
   {
      /// <summary>
      /// Creates a new instance of the specified class from within the script engine and returns it.
      /// </summary>
      /// <param name="engine"></param>
      /// <param name="className"></param>
      /// <returns></returns>
      public static object New(this ActiveScriptEngine engine, string className)
      {
         if (engine.ProgID.Equals(VBScript.ProgID, StringComparison.OrdinalIgnoreCase))
         {
            return engine.Evaluate("New " + className);
         }

         return engine.Evaluate("new " + className + "()");
      }

      public static T Evaluate<T>(this ActiveScriptEngine engine, string code)
      {
         return (T)Convert.ChangeType(engine.Evaluate(code), typeof(T));
      }
   }
}