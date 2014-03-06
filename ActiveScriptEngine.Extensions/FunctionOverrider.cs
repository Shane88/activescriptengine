namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Diagnostics.CodeAnalysis;
   using System.Globalization;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   public class FunctionOverrider
   {
      private Action<object[]> whenCalledAction;
      private Func<object[], object> whenCalledFunction;

      public void WhenCalled(Action<object[]> whenActionIsCalled)
      {
         this.whenCalledAction = whenActionIsCalled;
      }

      public void WhenCalled(Func<object[], object> whenFunctionIsCalled)
      {
         this.whenCalledFunction = whenFunctionIsCalled;
      }

      #region Func Overloads

      public void WhenCalled<T1>(Func<T1, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2>(Func<T1, T2, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3>(Func<T1, T2, T3, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4>(Func<T1, T2, T3, T4, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5>(Func<T1, T2, T3, T4, T5, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6>(Func<T1, T2, T3, T4, T5, T6, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7>(Func<T1, T2, T3, T4, T5, T6, T7, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8>(Func<T1, T2, T3, T4, T5, T6, T7, T8, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture), (T15)Convert.ChangeType(args[14], typeof(T15), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16>(Func<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, object> whenFunctionIsCalled) 
      {
         this.whenCalledFunction = (args) => whenFunctionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture), (T15)Convert.ChangeType(args[14], typeof(T15), CultureInfo.InvariantCulture), (T16)Convert.ChangeType(args[15], typeof(T16), CultureInfo.InvariantCulture));
      }

      #endregion Func Overloads
      
            #region Action Overloads

      public void WhenCalled<T1>(Action<T1> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2>(Action<T1, T2> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3>(Action<T1, T2, T3> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4>(Action<T1, T2, T3, T4> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5>(Action<T1, T2, T3, T4, T5> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6>(Action<T1, T2, T3, T4, T5, T6> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7>(Action<T1, T2, T3, T4, T5, T6, T7> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8>(Action<T1, T2, T3, T4, T5, T6, T7, T8> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture), (T15)Convert.ChangeType(args[14], typeof(T15), CultureInfo.InvariantCulture));
      }

      public void WhenCalled<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16>(Action<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16> whenActionIsCalled) 
      {
         this.whenCalledAction = (args) => whenActionIsCalled((T1)Convert.ChangeType(args[0], typeof(T1), CultureInfo.InvariantCulture), (T2)Convert.ChangeType(args[1], typeof(T2), CultureInfo.InvariantCulture), (T3)Convert.ChangeType(args[2], typeof(T3), CultureInfo.InvariantCulture), (T4)Convert.ChangeType(args[3], typeof(T4), CultureInfo.InvariantCulture), (T5)Convert.ChangeType(args[4], typeof(T5), CultureInfo.InvariantCulture), (T6)Convert.ChangeType(args[5], typeof(T6), CultureInfo.InvariantCulture), (T7)Convert.ChangeType(args[6], typeof(T7), CultureInfo.InvariantCulture), (T8)Convert.ChangeType(args[7], typeof(T8), CultureInfo.InvariantCulture), (T9)Convert.ChangeType(args[8], typeof(T9), CultureInfo.InvariantCulture), (T10)Convert.ChangeType(args[9], typeof(T10), CultureInfo.InvariantCulture), (T11)Convert.ChangeType(args[10], typeof(T11), CultureInfo.InvariantCulture), (T12)Convert.ChangeType(args[11], typeof(T12), CultureInfo.InvariantCulture), (T13)Convert.ChangeType(args[12], typeof(T13), CultureInfo.InvariantCulture), (T14)Convert.ChangeType(args[13], typeof(T14), CultureInfo.InvariantCulture), (T15)Convert.ChangeType(args[14], typeof(T15), CultureInfo.InvariantCulture), (T16)Convert.ChangeType(args[15], typeof(T16), CultureInfo.InvariantCulture));
      }

      #endregion Action Overloads
               
      [SuppressMessage("Microsoft.Design", "CA1043:UseIntegralOrStringArgumentForIndexers",
         Justification="This is not a typical indexer and is actually used to fake a default function in script")]
      public object this[params object[] args]
      {
         get
         {
            if (this.whenCalledAction != null)
            {
               this.whenCalledAction(args);
            }

            if (this.whenCalledFunction != null)
            {
               return this.whenCalledFunction(args);
            }

            return null;
         }
      }
   }
}