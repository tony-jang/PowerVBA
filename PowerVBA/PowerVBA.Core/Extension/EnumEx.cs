using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Extension
{
    public static class EnumEx
    {
        public static string GetDescription(this Enum Enum)
        {
            try
            {
                return ((DescriptionAttribute)Enum.GetType().GetField(Enum.ToString()).GetCustomAttributes(false).Where((i) => i.GetType() == typeof(DescriptionAttribute)).First()).Description;
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
