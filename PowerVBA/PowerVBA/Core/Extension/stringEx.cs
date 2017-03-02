using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Core.Extension
{
    static class stringEx
    {
        public static string Change(this string str, int startIndex, int Length, string text)
        {
            int index = 0;
            List<char> strChar = new List<char>();

            foreach (char c in str.ToCharArray())
            {
                if (!(startIndex <= index && startIndex + Length - 1 >= index))
                    strChar.Add(c);

                index++;
            }

            return new string(strChar.ToArray()).Insert(startIndex, text);
        }

        public static string[] SplitByNewLine(this string str)
        {
            return str.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        }
    }
}