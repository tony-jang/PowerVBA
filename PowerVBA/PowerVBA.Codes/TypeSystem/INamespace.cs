using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 해석된 네임스페이스를 나타냅니다.
    /// </summary>
    public interface INamespace : ISymbol, ICompilationProvider
    {
        /// <summary>
        /// 이 네임스페이스의 외부 닉네임을 가져옵니다. 일반 네임스페이스에 대해 빈 문자열을 반환합니다.
        /// </summary>
        string ExternAlias { get; }


        /// <summary>
        /// 이 네임스페이스의 전체 이름을 가져옵니다. (ex. VBA.Collection)
        /// </summary>
        string FullName { get; }

        /// <summary>
        /// 이 네임스페이스의 짧은 이름을 가져옵니다. (ex. Collection)
        /// </summary>
        new string Name { get; }


        /// <summary>
        /// 부모 네임스페이스를 가져옵니다.
        /// 만약 이 네임스페이스가 루트라면 null을 반환합니다.
        /// </summary>
        INamespace ParentNamespace { get; }


        /// <summary>
        /// 이 네임스페이스에 대해 모든 자식 네임스페이스를 가져옵니다.
        /// </summary>
        IEnumerable<INamespace> ChildNamespaces { get; }


        /// <summary>
        /// 이 네임스페이스의 타입들을 가져옵니다.
        /// </summary>
        IEnumerable<ITypeDefinition> Types { get; }


    }
}
