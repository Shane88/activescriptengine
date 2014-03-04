namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   public class FunctionOverrider
   {
      private Action<object[]> whenCalledAction;
      private Func<object[], object> whenCalledFunction;

      public void WhenCalled(Action<object[]> whenCalled)
      {
         this.whenCalledAction = whenCalled;
      }

      public void WhenCalled(Func<object[], object> whenCalled)
      {
         this.whenCalledFunction = whenCalled;
      }

      #region Func Overloads

      public void WhenCalled<TParam1>(Func<TParam1, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0]);
      }

      public void WhenCalled<TParam1, TParam2>(Func<TParam1, TParam2, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3>(Func<TParam1, TParam2, TParam3, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4>(Func<TParam1, TParam2, TParam3, TParam4, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13], (TParam15)args[14]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16, object> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13], (TParam15)args[14], (TParam16)args[15]);
      }

      #endregion Func Overloads
      
            #region Action Overloads

      public void WhenCalled<TParam1>(Action<TParam1> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0]);
      }

      public void WhenCalled<TParam1, TParam2>(Action<TParam1, TParam2> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3>(Action<TParam1, TParam2, TParam3> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4>(Action<TParam1, TParam2, TParam3, TParam4> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5>(Action<TParam1, TParam2, TParam3, TParam4, TParam5> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13], (TParam15)args[14]);
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)args[0], (TParam2)args[1], (TParam3)args[2], (TParam4)args[3], (TParam5)args[4], (TParam6)args[5], (TParam7)args[6], (TParam8)args[7], (TParam9)args[8], (TParam10)args[9], (TParam11)args[10], (TParam12)args[11], (TParam13)args[12], (TParam14)args[13], (TParam15)args[14], (TParam16)args[15]);
      }

      #endregion Action Overloads
         
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