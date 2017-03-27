using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
	/// 해석되지 않은 독립체를 나타냅니다.
	/// </summary>
	public interface IUnresolvedEntity : INamedElement, IHasAccessibility
    {
        /// <summary>
        /// 독립체 타입을 가져옵니다.
        /// </summary>
        SymbolKind SymbolKind { get; }

        /// <summary>
        /// Gets the complete entity region (including header+body)
        /// </summary>
        DomRegion Region { get; }

        /// <summary>
        /// Gets the entity body region.
        /// </summary>
        DomRegion BodyRegion { get; }


        /// <summary>
        /// Gets the declaring class.
        /// For members, this is the class that contains the member.
        /// For nested classes, this is the outer class. For top-level entities, this property returns null.
        /// </summary>
        IUnresolvedTypeDefinition DeclaringTypeDefinition { get; }

        /// <summary>
        /// 현재 독립체가 정적인지에 대한 여부를 가져옵니다.
        /// </summary>
        bool IsStatic { get; }

        /// <summary>
        /// 매크로나 컴파일러에 의해 만들어졌는지 여부를 가져옵니다.
        /// </summary>
        bool IsSynthetic { get; }
    }



    /// <summary>
    /// 해석된 독립체를 나타냅니다.
    /// </summary>
    public interface IEntity : ISymbol, ICompilationProvider, INamedElement, IHasAccessibility
    {
        /// <summary>
        /// 독립체의 짧은 이름을 가져옵니다.
        /// </summary>
        new string Name { get; }

        ITypeDefinition DeclaringTypeDefinition { get; }


        IType DeclaringType { get; }

        IAssembly ParentAssembly { get; }


        /// <summary>
        /// Gets the complete entity region (including header+body)
        /// </summary>
        DomRegion Region { get; }

        /// <summary>
        /// Gets the entity body region.
        /// </summary>
        DomRegion BodyRegion { get; }



        /// <summary>
        /// 현재 독립체가 정적인지에 대한 여부를 가져옵니다.
        /// </summary>
        bool IsStatic { get; }

        /// <summary>
        /// 매크로나 컴파일러에 의해 만들어졌는지 여부를 가져옵니다.
        /// </summary>
        bool IsSynthetic { get; }
    }
}
