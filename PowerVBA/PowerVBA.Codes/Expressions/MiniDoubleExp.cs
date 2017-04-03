using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniDoubleExp : IMiniExp
    {
        public MiniDoubleExp(double Double)
        {
            this.Double = Double;
        }

        public double Double { get; set; }
        
        public MiniStrExp ConvertToString()
        {
            return new MiniStrExp(Convert.ToString(Double));
        }
    }
}
