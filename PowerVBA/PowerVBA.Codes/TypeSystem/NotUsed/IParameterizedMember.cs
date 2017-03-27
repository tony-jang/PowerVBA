using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IUnresolvedParameterizedMember : IUnresolvedMember
    {
        IList<IUnresolvedParameter> Paremeters { get; }
    }
    /// <summary>
    /// 매개 변수화 된 멤버를 나타냅니다. (메소드 또는 프로퍼티)
    /// </summary>
    public interface IParameterizedMember : IMember
    {
        /// <summary>
        /// 파라미터들입니다.
        /// </summary>
        IList<IParameter> Parameters { get; }
    }
}
