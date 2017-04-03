using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniIntExp : IMiniExp
    {
        public MiniIntExp(int Integer)
        {
            this.Integer = Integer;
        }

        public int Integer { get; set; }

        public MiniStrExp ConvertToString()
        {
            return new MiniStrExp(Integer.ToString());
        }
    }
}
