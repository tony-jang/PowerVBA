using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    /// <summary>
    /// 현재 라인에 대한 코드 아이템을 가져옵니다.
    /// </summary>
    class LineCodeItem : CodeItemBase
    {
        public LineCodeItem(string FileName, (int, int) Segment) : base(FileName, Segment)
        { }
        public LineCodeItem((int, int) Segment) : base(string.Empty, Segment)
        { }
    }
}
