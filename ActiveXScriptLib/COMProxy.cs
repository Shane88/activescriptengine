namespace ActiveXScriptLib
{
   using System;
   using System.Collections;
   using System.Diagnostics;
   using System.Globalization;
   using System.Reflection;
   using System.Runtime.InteropServices;

   [DebuggerDisplay("ComProxy for {UnderlyingSystemType}")]
   public class ComProxy : IReflect
   {
      private object actualValue;

      public ComProxy(object value)
      {
         if (value == null)
         {
            throw new ArgumentNullException("value");
         }

         actualValue = value;
         UnderlyingSystemType = value.GetType();
      }

      public object Value
      {
         get
         {
            return actualValue;
         }
      }

      public Type UnderlyingSystemType { get; private set; }

      public FieldInfo GetField(string name, BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetField(name, bindingAttr);
      }

      public FieldInfo[] GetFields(BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetFields(bindingAttr);
      }

      public MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetMember(name, bindingAttr);
      }

      public MemberInfo[] GetMembers(BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetMembers(bindingAttr);
      }

      public MethodInfo GetMethod(string name, BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetMethod(name, bindingAttr);
      }

      public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
      {
         return UnderlyingSystemType.GetMethod(name, bindingAttr, binder, types, modifiers);
      }

      public MethodInfo[] GetMethods(BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetMethods(bindingAttr);
      }

      public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetProperties(bindingAttr);
      }

      public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
      {
         return UnderlyingSystemType.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
      }

      public PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
      {
         return UnderlyingSystemType.GetProperty(name, bindingAttr);
      }

      public object InvokeMember(
         string name,
         BindingFlags invokeAttr,
         Binder binder,
         object target,
         object[] args,
         ParameterModifier[] modifiers,
         CultureInfo culture,
         string[] namedParameters)
      {
         object returnValue = null;

         if (name == "[DISPID=-4]")
         {
            IEnumerable enumerator = actualValue as IEnumerable;

            if (enumerator != null)
            {
               //return new ComProxy(enumerator.GetEnumerator() as IEnumerator);
               return enumerator.GetEnumerator() as IEnumerator;
            }
         }
         else
         {
            returnValue =
               UnderlyingSystemType.InvokeMember(name, invokeAttr, binder, actualValue, args, modifiers, culture, namedParameters);
         }

         if (ShouldWrapObject(returnValue))
         {
            return new ComProxy(returnValue);
         }

         return returnValue;
      }

      public static bool ShouldWrapObject(object value)
      {
         return value != null && 
            !Marshal.IsComObject(value) &&
            !ComProxy.IsComVisibleDotNetObject(value) &&
            !(value is ComProxy);
      }

      public static bool IsComVisibleDotNetObject(object value)
      {
         bool? comVisible = null;

         Type type = value.GetType();

         if (value is IReflect)
         {
            return true;
         }

         var comVisibleAttributes = 
            type.GetCustomAttributes(typeof(ComVisibleAttribute), false) as ComVisibleAttribute[];

         if (comVisibleAttributes != null)
         {
            foreach (var comAttribute in comVisibleAttributes)
            {
               comVisible = comAttribute.Value;
               break;
            }
         }

         if (!comVisible.HasValue ||
            (comVisible.HasValue && !comVisible.Value))
         {
            comVisibleAttributes =
               type.Assembly.GetCustomAttributes(typeof(ComVisibleAttribute), false) as ComVisibleAttribute[];

            if (comVisibleAttributes != null)
            {
               foreach (var comAttribute in comVisibleAttributes)
               {
                  comVisible = comAttribute.Value;
                  break;
               }
            }
         }

         return comVisible ?? false;
      }
   }
}