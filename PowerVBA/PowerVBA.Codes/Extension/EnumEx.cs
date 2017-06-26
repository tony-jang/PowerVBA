using PowerVBA.Codes.Attributes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes.Extension
{
    public static class EnumEx
    {
        public static string GetDescription(this Enum Enum, string CultureName)
        {
            try
            {
                
                return ((BaseErrorAttribute)Enum.GetType()
                                                .GetField(Enum.ToString())
                                                .GetCustomAttributes(false)
                                                .Where((i) => i.GetType().BaseType == typeof(BaseErrorAttribute))
                                                .Where((i) => CultureName == ((BaseErrorAttribute)i).MessageCulture.Name)
                                                       .First()).ErrorMessage;
            }
            catch
            {
                return "";
            }
        }

        public static string GetValue(this Enum Enum)
        {
            try
            {
                return ((ValueAttribute)Enum.GetType()
                       .GetField(Enum.ToString())
                       .GetCustomAttributes(false)
                       .Where((i) => i.GetType() == typeof(ValueAttribute))
                       .First()).Value;
            }
            catch 
            {
                return "";
            }
        }

        
        public static int GetPriority(this Enum Enum)
        {
            try
            {
                return ((PriorityAttribute)Enum.GetType()
                   .GetField(Enum.ToString())
                   .GetCustomAttributes(false)
                   .Where((i) => i.GetType() == typeof(PriorityAttribute))
                   .First()).Priority;
            }
            catch (Exception)
            {
                return -1;
            }
            
        }

        public static bool ContainsAttribute(this Enum Enum, Type type)
        {
            return Enum.GetType()
                       .GetField(Enum.ToString())
                       .GetCustomAttributes(false)
                       .Where((i) => i.GetType() == type).FirstOrDefault() != null;
        }
        public static int GetReplaceCount(this Enum Enum)
        {
            try
            {
                var itm = ((CanReplaceAttribute)Enum.GetType()
                                                .GetField(Enum.ToString())
                                                .GetCustomAttributes(false)
                                                .Where((i) => i.GetType() == typeof(CanReplaceAttribute))
                                                       .FirstOrDefault());
                if (itm != null) return itm.ReplaceCount;
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Enum.ToString() + Environment.NewLine + ex.ToString());
                return 0;
            }
        }
    }
}
