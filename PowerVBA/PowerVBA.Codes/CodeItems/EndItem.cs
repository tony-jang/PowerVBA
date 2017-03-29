using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class EndItem : CodeItemBase
    {
        public EndItem(ClosingItem itm, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            ClosingItem = itm;
        }
        public EndItem(ClosingItem itm, (int, int) Segment) : base(string.Empty, Segment)
        {
            ClosingItem = itm;
        }

        public ClosingItem ClosingItem { get; set; }
    }
}
