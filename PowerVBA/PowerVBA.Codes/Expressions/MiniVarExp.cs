using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    class MiniVarExp : IMiniExp
    {
        public MiniVarExp(string VariableLocation)
        {
            this.VariableLocation = VariableLocation;
        }

        /// <summary>
        /// 변수가 위치한 곳을 나타냅니다.
        /// </summary>
        string VariableLocation { get; set; }
    }
}
