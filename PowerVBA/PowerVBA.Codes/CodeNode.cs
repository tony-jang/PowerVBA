using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public class CodeNode
    {
        private List<CodeNode> ChildNode;
        public CodeNode()
        {
            ChildNode = new List<CodeNode>();
        }

        public RangeInt Lines;

        private ISegment _Segment;
        public ISegment Segment => _Segment;

        private CodeType _type;
        public CodeType Type => _type;
    }
}
