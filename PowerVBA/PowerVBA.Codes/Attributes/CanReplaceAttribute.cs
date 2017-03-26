using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Attributes
{
    /// <summary>
    /// 오류에 포함된 %%를 몇개까지 바꿀수 있는지에 대한 여부입니다.
    /// </summary>
    class CanReplaceAttribute : Attribute
    {
        public int ReplaceCount { get; }
        public CanReplaceAttribute(int replaceCount)
        {
            ReplaceCount = replaceCount;
        }
    }
}
