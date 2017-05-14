using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Extension
{
    public static class ArrayEx
    {
        public static IEnumerable<T> SubArray<T>(this IEnumerable<T> list, int startIndex)
        {
            return list.Skip(2);
        }
        public static IEnumerable<T> SubArray<T>(this IEnumerable<T> list, int startIndex, int length)
        {
            return list.Skip(2).Take(length);
        }
    }
}
