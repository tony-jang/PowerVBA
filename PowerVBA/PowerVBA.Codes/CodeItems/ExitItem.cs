using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class ExitItem : CodeItemBase
    {
        public ExitItem(CanExitItems type, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            DeclaratorType = type;
        }
        public ExitItem(CanExitItems type, (int, int) Segment) : base(string.Empty, Segment)
        {
            DeclaratorType = type;
        }

        CanExitItems DeclaratorType { get; set; }
    }
}
