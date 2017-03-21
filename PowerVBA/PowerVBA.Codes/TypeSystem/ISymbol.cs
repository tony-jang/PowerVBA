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
    /// TypeSystem Symbol에 대한 인터페이스.
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
    }

    public interface ISymbolReference
    {
        ISymbol Resolve(ITypeResolveContext context);
    }
}
