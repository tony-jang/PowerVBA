using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IUnresolvedParameter
    {
        string Name { get; }


        ITypeReference Type { get; }

        /// <summary>
        /// 이 매개 변수가 VBA의 'ByRef' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsRef { get; }

        /// <summary>
        /// 이 매개 변수가 VBA의 'ParamArray' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsParams { get; }
        /// <summary>
        /// 이 매개 변수가 VBA의 'Optional' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsOptional { get; }

        IParameter CreateResolvedParameter(ITypeResolveContext context);
    }

    public interface IParameter : IVariable
    {
        /// <summary>
        /// 이 매개 변수가 VBA의 'ByRef' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsRef { get; }

        /// <summary>
        /// 이 매개 변수가 VBA의 'ParamArray' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsParams { get; }
        /// <summary>
        /// 이 매개 변수가 VBA의 'Optional' 매개 변수인지 여부를 가져옵니다.
        /// </summary>
        bool IsOptional { get; }


        /// <summary>
        /// 이 파라미터의 소유자를 가져옵니다. (매개 변수를 가지고 있는 멤버)
        /// </summary>
        IParameterizedMember Owner { get; }
    }
}
