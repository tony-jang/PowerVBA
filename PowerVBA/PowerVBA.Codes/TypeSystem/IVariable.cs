using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IVariable : ISymbol
    {
        /// <summary>
        /// 변수의 이름을 가져옵니다.
        /// </summary>
        new string Name { get; }
        
    }
}
