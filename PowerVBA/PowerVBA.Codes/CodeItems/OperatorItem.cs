using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class OperatorItem : CodeItemBase
    {
        public OperatorItem(string Operator, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            this.Operator = Operator;
        }

        public OperatorItem(string Operator, (int, int) Segment) : base(string.Empty, Segment)
        {
            this.Operator = Operator;
        }

        public string Operator { get; set; }
    }
}
