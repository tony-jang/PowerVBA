using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniBooleanExp : IMiniExp
    {
        public MiniBooleanExp(bool Boolean)
        {
            this.Boolean = Boolean;
        }

        public bool Boolean { get; set; }

        public MiniIntExp ConvertToInt()
        {
            if (Boolean)
                return new MiniIntExp(-1);
            else
                return new MiniIntExp(0);
        }
    }
}
