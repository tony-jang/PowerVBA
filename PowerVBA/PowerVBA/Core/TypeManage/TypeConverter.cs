using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.TypeManage
{
    static class TypeConverter
    {
        static TypeConverter()
        {
            TypeDatas = new List<Type>();
        }

        public static string WordCheck(string Word)
        {
            foreach (Type t in TypeDatas)
                if (t.Name.ToLower() == Word.ToLower())
                    return t.Name;

            return Word;
        }

        public static Type GetTypeByWord(string Word)
        {
            foreach (Type t in TypeDatas)
                if (t.Name.ToLower() == Word.ToLower())
                    return t;

            return null;
        }


        static List<Type> TypeDatas;
    }
}