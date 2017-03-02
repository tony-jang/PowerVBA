using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Core.CodeEdit.Parser;

namespace PowerVBA.Core.Project
{
    class SingleLineInfo : ILineInfo
    {
        public SingleLineInfo(int Length, int Offset, LineType lineType, 
            bool IsInClass = false, bool IsInEnum = false, bool IsInIf = false, 
            bool IsInSelect = false, bool IsInType = false)
        {
            this.Length = Length;
            this.Offset = Offset;
            this.LineType = lineType;

            this.IsInClass = IsInClass;
            this.IsInEnum = IsInEnum;
            this.IsInIf = IsInIf;
            this.IsInSelect = IsInSelect;
            this.IsInType = IsInType;
        }

        public bool IsInClass { get; set; }
        public bool IsInEnum { get; set; }
        public bool IsInIf { get; set; }
        public bool IsInSelect { get; set; }
        public bool IsInType { get; set; }


        public int Length { get; set; }
        
        public int Offset { get; set; }

        public LineType LineType { get; set; }
    }
}
