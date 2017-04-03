using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniOperatorExp : IMiniExp
    {
        public MiniOperatorExp(ExpOperator Operator)
        {
            this.Operator = Operator;
        }

        public ExpOperator Operator { get; set; }
    }
}
