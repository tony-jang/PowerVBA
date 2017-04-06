using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class LabelItem : CodeItemBase
    {
        public LabelItem(string FileName,string LabelName, (int, int) Segment) : base(FileName, Segment)
        {
            this.LabelName = LabelName;
        }
        public string LabelName { get; set; }
    }
}
