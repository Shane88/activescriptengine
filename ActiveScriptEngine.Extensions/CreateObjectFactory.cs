namespace ActiveXScriptLib.Extensions
{
   using System;
   using System.Collections.Generic;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   public class CreateObjectFactory
   {
      private Dictionary<string, object> objectFactories = new Dictionary<string, object>();

      // TODO: Should consider making this take a factory so it can be deferred like the rest of the api.
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

   public static class ActiveScriptEngineBuilderCreateObjectFactoryExtensions
   {
      public static ActiveScriptEngineBuilder AddCreateObjectHook(this ActiveScriptEngineBuilder builder, Action<CreateObjectFactory> factory)
      {
         builder.Configure(engine => ConfigureCreateObjectHookImpl(engine, factory));
         return builder;
      }

      private static void ConfigureCreateObjectHookImpl(ActiveScriptEngine engine, Action<CreateObjectFactory> factory)
      {
         CreateObjectFactory createObjectFactory;

         if (engine.ScriptHasHostObject("CreateObject"))
         {
            object createObject = engine.GetScriptObject("CreateObject");

            var existingFactory = createObject as CreateObjectFactory;

            if (existingFactory != null)
            {
               createObjectFactory = existingFactory;
            }
            else
            {
               string errorMessage = string.Format(
                  "When attempting to configure CreateObjectFactory and existing object was found to already be added to the script engine" +
                  "with the name \"CreateObject\" but it was of type {0} instead of the expected type of {1}",
                  createObject == null ? "Null" : createObject.GetType().FullName,
                  typeof(CreateObjectFactory).FullName);

               throw new InvalidOperationException(errorMessage);
            }
         }
         else
         {
            createObjectFactory = new CreateObjectFactory();

            engine.AddObject("CreateObject", createObjectFactory);
         }

         factory(createObjectFactory);
      }
   }
}
