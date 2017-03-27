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


        /// <summary>
        /// Gets the declaration region of the variable.
        /// </summary>
        DomRegion Region { get; }

        /// <summary>
        /// 변수의 형식을 가져옵니다.
        /// </summary>
        IType Type { get; }

        /// <summary>
        /// 이 변수가 상수인지에 대한 여부를 가져옵니다.
        /// </summary>
        bool IsConst { get; }

        /// <summary>
        /// 이 필드가 상수이면 값을 검색합니다.
        /// 매개 변수의 경우 이것이 기본값입니다.
        /// </summary>
        object ConstantValue { get; }

    }
}
