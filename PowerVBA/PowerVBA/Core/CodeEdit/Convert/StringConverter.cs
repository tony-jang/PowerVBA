using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Convert
{
    static class StringConverter
    {
        /// <summary>
        /// 선언 타입 (Sub, Function, Type 등) 들을 원래 포맷으로 변경해줍니다.
        /// </summary>
        /// <param name="Type"></param>
        /// <returns></returns>
        public static string ConvertType(string Type)
        {
            switch (Type.ToLower())
            {
                case "sub":
                    return "Sub";
                case "function":
                    return "Function";
                case "type":
                    return "Type";
                case "enum":
                    return "Enum";
                case "select":
                    return "Select";
                default: return Type;
            }
        }

        /// <summary>
        /// 엑세서 (Public, Private, Dim 등) 들을 원래 포맷으로 변경해줍니다.
        /// </summary>
        /// <param name="Accessor"></param>
        /// <returns></returns>
        public static string ConvertAccessor(string Accessor)
        {
            switch (Accessor.ToLower())
            {
                case "public":
                    return "Public";
                case "private":
                    return "Private";
                case "dim":
                    return "Dim";
                default: return Accessor;
            }
        }
    }
}
