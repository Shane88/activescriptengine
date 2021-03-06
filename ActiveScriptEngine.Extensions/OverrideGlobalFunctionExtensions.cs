﻿namespace ActiveXScriptLib.Extensions
{
   using System;

   public static class OverrideGlobalFunctionExtensions
   {
      public static void OverrideGlobalFunction(this ActiveScriptEngine engine, string functionName, Action<object[]> action)
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }

         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction(this ActiveScriptEngine engine, string functionName, Func<object[], object> func)
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }

         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1>(this ActiveScriptEngine engine, string functionName, Func<T1, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16>(this ActiveScriptEngine engine, string functionName, Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, object> func) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
            
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(func);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1>(this ActiveScriptEngine engine, string functionName, Action<T1> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2>(this ActiveScriptEngine engine, string functionName, Action<T1, T2> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }

      public static void OverrideGlobalFunction<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16>(this ActiveScriptEngine engine, string functionName, Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16> action) 
      {
         if (engine == null)
         {
            throw new ArgumentNullException("engine");
         }
               
         var overrider = new FunctionOverrider();
         overrider.WhenCalled(action);
         engine.AddObject(functionName, overrider);
      }
   }
}
