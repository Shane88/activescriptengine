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
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)));
      }

      public void WhenCalled<TParam1, TParam2>(Func<TParam1, TParam2, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3>(Func<TParam1, TParam2, TParam3, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4>(Func<TParam1, TParam2, TParam3, TParam4, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)), (TParam15)Convert.ChangeType(args[14], typeof(TParam15)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16>(Func<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16, object> whenCalled) 
      {
         this.whenCalledFunction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)), (TParam15)Convert.ChangeType(args[14], typeof(TParam15)), (TParam16)Convert.ChangeType(args[15], typeof(TParam16)));
      }

      #endregion Func Overloads
      
            #region Action Overloads

      public void WhenCalled<TParam1>(Action<TParam1> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)));
      }

      public void WhenCalled<TParam1, TParam2>(Action<TParam1, TParam2> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3>(Action<TParam1, TParam2, TParam3> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4>(Action<TParam1, TParam2, TParam3, TParam4> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5>(Action<TParam1, TParam2, TParam3, TParam4, TParam5> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)), (TParam15)Convert.ChangeType(args[14], typeof(TParam15)));
      }

      public void WhenCalled<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16>(Action<TParam1, TParam2, TParam3, TParam4, TParam5, TParam6, TParam7, TParam8, TParam9, TParam10, TParam11, TParam12, TParam13, TParam14, TParam15, TParam16> whenCalled) 
      {
         this.whenCalledAction = (args) => whenCalled((TParam1)Convert.ChangeType(args[0], typeof(TParam1)), (TParam2)Convert.ChangeType(args[1], typeof(TParam2)), (TParam3)Convert.ChangeType(args[2], typeof(TParam3)), (TParam4)Convert.ChangeType(args[3], typeof(TParam4)), (TParam5)Convert.ChangeType(args[4], typeof(TParam5)), (TParam6)Convert.ChangeType(args[5], typeof(TParam6)), (TParam7)Convert.ChangeType(args[6], typeof(TParam7)), (TParam8)Convert.ChangeType(args[7], typeof(TParam8)), (TParam9)Convert.ChangeType(args[8], typeof(TParam9)), (TParam10)Convert.ChangeType(args[9], typeof(TParam10)), (TParam11)Convert.ChangeType(args[10], typeof(TParam11)), (TParam12)Convert.ChangeType(args[11], typeof(TParam12)), (TParam13)Convert.ChangeType(args[12], typeof(TParam13)), (TParam14)Convert.ChangeType(args[13], typeof(TParam14)), (TParam15)Convert.ChangeType(args[14], typeof(TParam15)), (TParam16)Convert.ChangeType(args[15], typeof(TParam16)));
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