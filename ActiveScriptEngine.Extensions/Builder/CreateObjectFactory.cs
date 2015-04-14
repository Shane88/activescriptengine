namespace ActiveXScriptLib.Extensions.Builder
{
   using System;
   using System.Collections.Generic;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   public class CreateObjectFactory
   {
      internal Dictionary<string, Func<object>> objectFactories = new Dictionary<string, Func<object>>(StringComparer.OrdinalIgnoreCase);
      internal Dictionary<string, object> builtObjects = new Dictionary<string, object>();

      public InterceptWith Intercept(string name)
      {
         return new InterceptWith(this, name);
      }

      public object this[string name]
      {
         get
         {
            if (builtObjects.ContainsKey(name))
            {
               return builtObjects[name];
            }

            if (objectFactories.ContainsKey(name))
            {
               // On first request of the object, build it.
               Func<object> factory = objectFactories[name];

               object instance = factory();

               // Cache built objects in another dictionary.
               builtObjects.Add(name, instance);

               // Don't need the factory any more.
               objectFactories.Remove(name);

               return instance;
            }

            return Activator.CreateInstance(Type.GetTypeFromProgID(name));
         }
      }
   }

   public class InterceptWith
   {
      private readonly CreateObjectFactory _factory;
      private readonly string _name;

      public InterceptWith(CreateObjectFactory factory, string name)
      {
         _factory = factory;
         _name = name;
      }

      public CreateObjectFactory With<T>() where T : class, new()
      {
         _factory.objectFactories.Add(_name, () => new T());
         return _factory;
      }

      public CreateObjectFactory With<T>(T value)
      {
         _factory.objectFactories.Add(_name, () => value);
         return _factory;
      }

      public CreateObjectFactory With<T>(Action<T> configurationAction) where T : class, new()
      {
         Func<object> factory = () =>
         {
            T instance = new T();
            configurationAction(instance);
            return instance;
         };

         _factory.objectFactories.Add(_name, factory);
         return _factory;
      }

      public CreateObjectFactory With<T>(Func<T> objectFactory) where T : class
      {
         _factory.objectFactories.Add(_name, objectFactory);
         return _factory;
      }
   }

   public static class ActiveScriptEngineBuilderCreateObjectFactoryExtensions
   {
      public static ActiveScriptEngineBuilder InterceptCreateObject(this ActiveScriptEngineBuilder builder, Action<CreateObjectFactory> factory)
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
