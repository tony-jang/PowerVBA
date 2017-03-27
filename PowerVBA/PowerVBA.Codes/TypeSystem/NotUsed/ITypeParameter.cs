using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 기본 클래스/메소드의 타입 파라미터입니다.
    /// </summary>
    public interface IUnresolvedTypeParameter : INamedElement
    {
        /// <summary>
        /// 파라미터의 소유자 타입을 가져옵니다.
        /// </summary>
        SymbolKind OwnerType { get; }

        /// <summary>
        /// 소유 메소드/클래스의 타입 매개 변수 목록에서 
        /// 형식 매개 변수의 인덱스를 가져옵니다.
        /// </summary>
        int Index { get; }

        /// <summary>
        /// Gets the region where the type parameter is defined.
        /// </summary>
        DomRegion Region { get; }

        /// <summary>
		/// 파라미터의 변화를 가져옵니다.
		/// </summary>
		VarianceModifier Variance { get; }



        ITypeParameter CreateResolvedTypeParameter(ITypeResolveContext context);
    }



    /// <summary>
	/// 기본 클래스/메소드의 타입 파라미터입니다.
	/// </summary>
	public interface ITypeParameter : IType, ISymbol
    {
        /// <summary>
        /// 파라미터의 소유자 타입을 가져옵니다.
        /// </summary>
        /// <returns>SymbolKind.TypeDefinition or SymbolKind.Method</returns>
        SymbolKind OwnerType { get; }

        /// <summary>
        /// 소유하고 있는 메소드/클래스를 가져옵니다.
        /// 이 속성은 null을 반환 할 수 있습니다.
        /// </summary>
        IEntity Owner { get; }

        /// <summary>
        /// Gets the region where the type parameter is defined.
        /// </summary>
        DomRegion Region { get; }

        /// <summary>
        /// 소유 메소드/클래스의 형식 파라미터 목록에서 형식 파라미터의 인덱스를 가져옵니다.
        /// </summary>
        int Index { get; }

        /// <summary>
        /// 형식 매개 변수의 이름을 가져옵니다.
        /// </summary>
        new string Name { get; }

        /// <summary>
        /// 이 타입의 파라미터의 분산을 가져옵니다.
        /// </summary>
        VarianceModifier Variance { get; }

        /// <summary>
        /// 이 타입의 파라미터의 유효한 기본 클래스를 가져옵니다.
        /// </summary>
        IType EffectiveBaseClass { get; }

        /// <summary>
        /// 이 형태 파라미터의 유효한 인터페이스 세트를 취득합니다.
        /// </summary>
        ICollection<IType> EffectiveInterfaceSet { get; }

        /// <summary>
        /// type 파라미터에 new() 제약 조건이 있는지 여부를 가져옵니다.
        /// </summary>
        bool HasDefaultConstructorConstraint { get; }

        /// <summary>
        /// 형식 매개 변수에 'class' 제약 조건이 있는지 여부를 가져옵니다.
        /// </summary>
        bool HasReferenceTypeConstraint { get; }

        /// <summary>
        /// 형식 매개 변수에 'struct' 제약 조건이 있는지 여부를 가져옵니다.
        /// </summary>
        bool HasValueTypeConstraint { get; }
    }



    /// <summary>
    /// 타입 파라미터의 분산을 나타냅니다.
    /// </summary>
    public enum VarianceModifier : byte
    {
        /// <summary>
        /// Type 파라미터는 변형이 아닙니다.
        /// </summary>
        Invariant,
        /// <summary>
        /// 형식 매개 변수는 같이 변합니다.
        /// </summary>
        Covariant,
        /// <summary>
        /// 형식 매개 변수는 반변수입니다.
        /// </summary>
        Contravariant
    };

}
