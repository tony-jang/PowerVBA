using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class AccessorItem : CodeItemBase
    {
        public AccessorItem(Accessor Accessor, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            this.Accessor = Accessor;
        }

        public AccessorItem(Accessor Accessor, (int,int) Segment) : base(string.Empty, Segment)
        {
            this.Accessor = Accessor;
        }

        public Accessor Accessor { get; set; }
        public string AccessorStr { get => Accessor.ToString(); }
        
    }
}
