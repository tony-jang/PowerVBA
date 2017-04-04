using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class KeywordItem : CodeItemBase
    {
        public KeywordItem(Keywords Keyword, string FileName, (int,int) Segment) : base(FileName, Segment)
        {
            this.Keyword = Keyword;
        }
        
        public Keywords Keyword { get; set; }

    }
}
