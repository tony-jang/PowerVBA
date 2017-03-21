using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// IType 참조 형식을 해석합니다.
    /// </summary>
    public interface ITypeReference
    {
        /// <summary>
        /// 해당 타입 참조 형식을 해석합니다.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        IType Resolve(ITypeResolveContext context);
    }

    public interface ITypeResolveContext : ICompilationProvider
    {
        /// <summary>
        /// 현재 어셈블리를 가져옵니다.
        /// </summary>
        IAssembly CurrentAssembly { get; }

        /// <summary>
        /// 현재 타입 정의를 가져옵니다.
        /// </summary>
        ITypeDefinition CurrentTypeDefinition { get; }

        /// <summary>
        /// 현재 멤버를 가져옵니다.
        /// </summary>
        IMember CurrentMember{ get; }

        ITypeResolveContext WithCurrentTypeDefinition(ITypeDefinition typeDefinition);
        ITypeResolveContext WithCurrentMemeber(IMember member);
    }
}
