using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class UnknownItem : CodeItemBase
    {
        public UnknownItem(string FileName, (int, int) Segment) : base(FileName, Segment)
        {
        }
        public UnknownItem((int, int) Segment) : base(string.Empty, Segment)
        {
        }
    }
}
