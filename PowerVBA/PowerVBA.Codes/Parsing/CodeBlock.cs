using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Parsing
{
    public class CodeBlock
    {
        public CodeBlock(LinePoint start, LinePoint end)
        {
            this.Start = start;
            this.End = end;
        }

        public CodeBlock (LinePoint start)
        {
            this.Start = start;
        }

        //public bool IsContains(LinePoint linePoint)
        //{
        //    if (!linePoint.IsOverlapped(Start) && !linePoint.IsOverlapped(End))
        //    {
        //        if (linePoint.)
        //        return true;
        //    }
        //}

        public List<CodeBlock> Childrens;

        public BlockType Type { get; set; }
        public LinePoint Start { get; set; }
        public LinePoint End { get; set; }
    }
}
