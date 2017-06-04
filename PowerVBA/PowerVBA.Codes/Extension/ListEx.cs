using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Extension
{
    public static class ListEx
    {
        /// <summary>
        /// 데이터를 확인한 뒤 존재하지 않으면 추가합니다.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="data">추가할 데이터입니다.</param>
        public static void CheckAdd<T>(this List<T> list, T data)
        {
            if (!list.Contains(data))
                list.Add(data);
        }
    }
}
