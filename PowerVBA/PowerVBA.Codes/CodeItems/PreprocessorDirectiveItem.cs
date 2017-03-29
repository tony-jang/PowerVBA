using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class PreprocessorDirectiveItem : CodeItemBase
    {
        public PreprocessorDirectiveItem(string FileName, (int, int) Segment) : base(FileName, Segment)
        {
        }
    }
}
