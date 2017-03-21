using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 해석되지 않은 어셈블리를 나타냅니다.
    /// </summary>
    public interface IUnresolvedAssembly : IAssemblyReference
    {
        /// <summary>
        /// 어셈블리 이름을 가져옵니다. (짧은 이름)
        /// </summary>
        string AssemblyName { get; }

        /// <summary>
        /// 어셈블리의 전체 이름을 가져옵니다. (Public Key, 토큰 등.. 포함)
        /// </summary>
        string FullAssemblyName { get; }


        /// <summary>
        /// 어셈블리 위치가 있는 경로를 가져옵니다.
        /// (프로젝트의 경우 출력 경로와 동일합니다.)
        /// </summary>
        string Location { get; }


        /// <summary>
        /// 어셈블리에서 중첩되지 않은 형식을 모두 가져옵니다.
        /// </summary>
        IEnumerable<IUnresolvedTypeDefinition> TopLevelTypeDefintions { get; }

    }

    public interface IAssemblyReference
    {
        /// <summary>
        /// 해당 어셈블리를 해석합니다.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        IAssembly Resolve(ITypeResolveContext context);
    }


    public interface IAssembly : ICompilationProvider
    {
        IUnresolvedAssembly UneresolvedAssembly { get; }

        /// <summary>
        /// 이 어셈블리가 컴파일의 주 어셈블리인지 여부를 가져옵니다.
        /// </summary>
        bool IsMainAssembly { get; }

        /// <summary>
        /// 어셈블리 이름을 가져옵니다. (짧은 이름)
        /// </summary>
        string AssemblyName { get; }

        /// <summary>
        /// 어셈블리의 전체 이름을 가져옵니다. (Public Key, 토큰 등.. 포함)
        /// </summary>
        string FullAssemblyName { get; }


        /// <summary>
        /// 이 어셈블리의 내부가 지정된 어셈블리에 표시되는지 여부를 가져옵니다.
        /// </summary>
        bool InternalsVisibleTo(IAssembly assembly);

        /// <summary>
        /// 이 어셈블리의 최상위 네임스페이스를 가져옵니다.
        /// </summary>
        /// <remarks>
        /// 이것은 항상 이름이없는 네임 스페이스입니다. 'root namespace' 프로젝트 설정과 관련이 없습니다.
        /// </remarks>
        INamespace RootNameSpace { get; }


        ITypeDefinition GetTypeDefinition(TopLevelTypeName topLevelTypeName);

        /// <summary>
        /// 어셈블리 내에 있는 모든 내포되지 않은 타입들을 가져옵니다.
        /// </summary>
        IEnumerable<ITypeDefinition> TopLevelTypeDefinitions { get; }
    }
}
