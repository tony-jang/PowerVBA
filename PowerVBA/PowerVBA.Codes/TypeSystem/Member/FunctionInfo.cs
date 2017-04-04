using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public class FunctionInfo : MethodInfo
    {
        public FunctionInfo(string Name, IType ReturnType, Accessor Accessor) : base(Name, Accessor)
        {
            this.ReturnType = ReturnType;
        }

        /// <summary>
        /// 반환 값을 나타냅니다.
        /// </summary>
        public IType ReturnType;
    }
}
