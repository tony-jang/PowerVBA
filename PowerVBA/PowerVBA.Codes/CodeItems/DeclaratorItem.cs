using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{

    
    class DeclaratorItem : CodeItemBase
    {
        public DeclaratorItem(DeclaratorType type, (int, int) Segment) : base(string.Empty, Segment)
        {
            DeclaratorType = type;
        }
        public DeclaratorItem(DeclaratorType type, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            DeclaratorType = type;
        }
        public DeclaratorType DeclaratorType { get; }
        
        public new string StrDescription { get => DeclaratorType.ToString(); }

    }
}
