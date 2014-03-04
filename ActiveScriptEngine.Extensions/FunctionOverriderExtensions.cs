namespace ActiveXScriptLib.Extensions
{
   using System;

   public static class FunctionOverriderExtensions
   {
      public static void OverrideGlobalFunction(this ActiveScriptEngine engine, string functionName, Action<object[]> action)
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction(this ActiveScriptEngine engine, string functionName, Func<object[], object> func)
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

         public static void OverrideGlobalFunction<TParam1>(this ActiveScriptEngine engine, string functionName, Func<TParam1, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(this ActiveScriptEngine engine, string functionName, Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, object> func) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

   
               public static void OverrideGlobalFunction<TParam1>(this ActiveScriptEngine engine, string functionName, Action<TParam1> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(this ActiveScriptEngine engine, string functionName, Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10> action) 
      {
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

   


   }
}