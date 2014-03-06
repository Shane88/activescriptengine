namespace ActiveXScriptLib
{
   using System;
   using System.Runtime.Serialization;

   [Serializable]
   public class InvalidActiveScriptClassException : Exception
   {
      public InvalidActiveScriptClassException()
      {
      }

      public InvalidActiveScriptClassException(string message)
         : base(message)
      {
      }

      public InvalidActiveScriptClassException(string message, Exception inner)
         : base(message, inner)
      {
      }

      protected InvalidActiveScriptClassException(
       SerializationInfo info,
       StreamingContext context)
         : base(info, context)
      {
      }
   }
}