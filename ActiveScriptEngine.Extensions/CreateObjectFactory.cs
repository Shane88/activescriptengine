namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Collections.Generic;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   public class CreateObjectFactory
   {
      private Dictionary<string, object> objectFactories = new Dictionary<string, object>();

      public CreateObjectFactory AddHook(string name, object value)
      {
         objectFactories.Add(name, value);
         return this;
      }

      public object this[string name]
      {
         get
         {
            if (objectFactories.ContainsKey(name))
            {
               return objectFactories[name];
            }

            return Activator.CreateInstance(Type.GetTypeFromProgID(name));
         }
      }
   }
}
