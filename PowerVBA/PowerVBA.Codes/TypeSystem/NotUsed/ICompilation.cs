using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 컴파일에 관련된 항목들을 나타냅니다.
    /// </summary>
    public interface ICompilation
    {
        /// <summary>
        /// 현재 어셈블리를 가져옵니다.
        /// </summary>
        IAssembly MainAssembly { get; }


        /// <summary>
        /// 이 컴파일을 지정하고 현재 어셈블리 또는 독립체를 지정하지 않는 유형 확인 컨텍스트를 가져옵니다.
        /// </summary>
        ITypeResolveContext TypeResolveContext { get; }


        /// <summary>
        /// 이 컴파일 내의 모든 어셈블리들을 가져옵니다.
        /// </summary>
        IList<IAssembly> Assemblies { get; }


        /// <summary>
        /// 참조된 어셈블리들을 가져옵니다.
        /// 이 리스트는 메인 어셈블리에 포함되지 않습니다.
        /// </summary>
        IList<IAssembly> ReferencedAssemblies { get; }


        /// <summary>
        /// 이 컴파일의 루트 네임스페이스를 가져옵니다.
        /// </summary>
        INamespace RootNamespace { get; }

        /// <summary>
        /// 지정된 extern 가명 루트 네임스페이스를 가져옵니다.
        /// </summary>
        /// <param name="alias"></param>
        /// <returns></returns>
        INamespace GetNamespaceForExternAlias(string alias);

        /// <summary>
        /// 알려진 타입을 찾습니다.
        /// </summary>
        /// <param name="typeCode"></param>
        /// <returns></returns>
        IType FindType(KnownTypeCode typeCode);


        /// <summary>
        /// 컴파일되고 있는 언어의 이름 비교자를 가져옵니다.
        /// INamespace.GetTypeDefinition 메소드에 사용되는 문자열 비교자 입니다.
        /// </summary>
        StringComparer NameComparer { get; }

        //ISolutionSnapshot SolutionSnapshot { get; }


        //CacheManager CacheManager { get; }
    }

    public interface ICompilationProvider
    {

        /// <summary>
        /// 부모 컴파일을 가져옵니다.
        /// 이 속성은 절대로 null을 반환하지 않습니다.
        /// </summary>
        ICompilation Compilation { get; }
    }
}
