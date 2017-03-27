using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public enum SymbolKind : byte
    {
        None,
        TypeDefinition,
        Field,
        Property,
        Event,
        Method,
        Operator,
        Accessor,
        Variable,
        Parameter,
        TypeParameter
    }
    /// <summary>
    /// 타입 시스템 Symbol에 대한 인터페이스입니다.
    /// </summary>
    public interface ISymbol
    {
        /// <summary>
        /// 이 속성은 어떤 종류의 Symbol인지 지정하는 enum을 반환합니다.
        /// </summary>
        SymbolKind SymbolKind { get; }

        /// <summary>
        /// Symbol의 간단한 이름을 가져옵니다.
        /// </summary>
        string Name { get; }


        /// <summary>
        /// 다른 컴파일에서 이 Symbol을 재발견하는 데 사용 가능한 심볼 참조를 작성합니다.
        /// </summary>
        /// <returns></returns>
        ISymbolReference ToReference();
    }

    public interface ISymbolReference
    {
        ISymbol Resolve(ITypeResolveContext context);
    }
}
