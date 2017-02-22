using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Convert
{
    static class StringConverter
    {
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

        public static string ConvertAccessor(string Accessor)
        {
            switch (Accessor.ToLower())
            {
                case "public":
                    return "Public";
                case "private":
                    return "Private";
                default: return Accessor;
            }
        }
    }
}
