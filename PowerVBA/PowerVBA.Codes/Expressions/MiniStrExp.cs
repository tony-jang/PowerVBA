using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniStrExp : IMiniExp
    {
        public MiniStrExp(string String)
        {
            this.String = String;
        }

        public string String { get; set; }


        public MiniIntExp ConvertToInt(out bool Error)
        {
            int i;
            Error = false;
            if (int.TryParse(String, out i))
            {
                return new MiniIntExp(i);
            }
            else
            {
                Error = false;
                return new MiniIntExp(int.MinValue);
            }
            
        }
    }
}
