using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class CommentItem : CodeItemBase
    {
        public CommentItem(string FileName, (int, int) Segment) : base(FileName, Segment)
        {
        }
        public CommentItem((int, int) Segment) : base(string.Empty, Segment)
        {
        }
    }
}
