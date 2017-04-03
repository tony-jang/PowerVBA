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


        public MiniIntExp ConvertToInt()
        {
            int i;
            if (int.TryParse(String, out i))
            {
                return new MiniIntExp(i);
            }
            else { throw new Exception($"'{String}'식 은 Int로 변환할 수 없습니다."); }
            
        }
    }
}
