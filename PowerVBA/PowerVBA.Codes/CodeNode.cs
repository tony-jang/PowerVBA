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
        
        public CodeNode()
        {
            _ChildNode = new List<CodeData>();
        }

        //private RangeInt _Lines;
        //public RangeInt Lines => _Lines;

        private List<CodeData> _ChildNode;
        public  List<CodeData> ChildNode => _ChildNode;
        
        
    }
}
