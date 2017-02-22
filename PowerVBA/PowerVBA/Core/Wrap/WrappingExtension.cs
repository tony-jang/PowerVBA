using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Wrap
{
    static class WrappingExtension
    {
        static Dictionary<object, object> cache =
            new Dictionary<object, object>();

        private static TAttribute GetAttribute<TAttribute>(this Type type)
            where TAttribute : Attribute
        {
            return type.GetCustomAttribute<TAttribute>();
        }

        private static bool HasAttribute<TAttribute>(this Type type)
            where TAttribute : Attribute
        {
            return type.GetAttribute<TAttribute>() != null;
        }

        public static TWrappingClass GetWrapping<TWrappingClass>(this object obj)
            where TWrappingClass : IWrappingClass
        {
            if (!cache.ContainsKey(obj))
            {
                string typeName = Microsoft.VisualBasic.Information.TypeName(obj);

                Type wrappingType =
                    Assembly.GetExecutingAssembly()
                        .GetTypes()
                        .Where(t => typeof(IWrappingClass).IsAssignableFrom(t))
                        .Where(t => t.HasAttribute<WrappedAttribute>())
                        .FirstOrDefault(t => t.GetAttribute<WrappedAttribute>().TargetType.Name == typeName);

                cache[obj] = (TWrappingClass)Activator.CreateInstance(wrappingType, new object[] { obj });
            }

            return (TWrappingClass)cache[obj];
        }
    }
}
