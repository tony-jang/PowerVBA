using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class TypeItem : CodeItemBase
    {
        public TypeItem(string TypeName, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            this.TypeName = TypeName;
        }
        public TypeItem(string TypeName, (int, int) Segment) : base("", Segment)
        {
            this.TypeName = TypeName;
        }

        public string TypeName { get; set; }

        /// <summary>
        /// 해당 CodeInfo로 부터 타입 값을 받아옵니다.
        /// </summary>
        /// <param name="ci"></param>
        /// <returns></returns>
        public IType ReturnType(CodeInfo ci)
        {
            return null;
        }
    }
}
