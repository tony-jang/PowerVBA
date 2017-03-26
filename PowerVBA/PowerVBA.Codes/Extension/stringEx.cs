using PowerVBA.Codes.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Extension
{
    public static class stringEx
    {
        public static bool IsReservedKeyWords(this string str)
        {
            return Resources.예약어.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Select((i)=> i.ToLower()).ToList().Contains(str.ToLower());
        }
    }
}
