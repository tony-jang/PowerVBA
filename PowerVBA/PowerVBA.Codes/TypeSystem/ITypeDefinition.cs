using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 해석되지 않은 클래스/Enum/Interface/Struct/Delegate를 나타냅니다.
    /// </summary>
    public interface IUnresolvedTypeDefinition : ITypeReference, IUnresolvedEntity
    {
        TypeKind Kind { get; }

        FullTypeName FullTypeName { get; }

        IList<ITypeReference> BaseTypes { get; }
        IList<IUnresolvedTypeParameter> TypeParameters { get; }

        IList<IUnresolvedTypeDefinition> NestedTypes { get; }
        IList<IUnresolvedMember> Members { get; }

        IEnumerable<IUnresolvedMethod> Methods { get; }
        IEnumerable<IUnresolvedProperty> Properties { get; }
        //IEnumerable<IUnresolvedField> Fields { get; }
        //IEnumerable<IUnresolvedEvent> Events { get; }

        
        /// <summary>
        /// 이 해석되지 않은 타입 정의가 기본 생성자를 추가하는지에 대한 여부를 가져옵니다.
        /// if no other constructor is present.
        /// </summary>
        bool AddDefaultConstructorIfRequired { get; }

        /// <summary>
        /// 이 해석되지 않은 정의에 대응하는 문맥으로부터 해결 된 형태 정의를 검색합니다.
        /// type definition.
        /// </summary>
        new IType Resolve(ITypeResolveContext context);

        /// <summary>
        /// 이 메서드는 언어 별 요소를 형식 확인 컨텍스트에 추가하는 데 사용합니다.
        /// </summary>
        ITypeResolveContext CreateResolveContext(ITypeResolveContext parentContext);
    }

    /// <summary>
    /// 클래스/Enum/Interface/Struct/Delegate를 나타냅니다.
    /// </summary>
    public interface ITypeDefinition : IType, IEntity
    {
        /// <summary>
        /// Returns all parts that contribute to this type definition.
        /// Non-partial classes have a single part that represents the whole class.
        /// </summary>
        IList<IUnresolvedTypeDefinition> Parts { get; }

        IList<ITypeParameter> TypeParameters { get; }

        IList<ITypeDefinition> NestedTypes { get; }
        IList<IMember> Members { get; }

        IEnumerable<IField> Fields { get; }
        IEnumerable<IMethod> Methods { get; }
        IEnumerable<IProperty> Properties { get; }
        //IEnumerable<IEvent> Events { get; }

        /// <summary>
        /// 이 형식 정의에 대해 알려진 형식 코드를 가져옵니다.
        /// </summary>
        KnownTypeCode KnownTypeCode { get; }

        /// <summary>
        /// Enum의 경우 : 기본 유형 반환 /
        /// 다른 모든 타입의 경우 : <see cref="SpecialType.UnknownType"/> 반환.
        /// </summary>
        IType EnumUnderlyingType { get; }

        /// <summary>
        /// Gets the full name of this type.
        /// </summary>
        /// <remarks>
        /// DEBUG: 테스트로 안 써보고 있음 사용하게 되면 다시 정의 할 것.
        /// </remarks>
        //FullTypeName FullTypeName { get; }

        /// <summary>
        /// 선언 타입을 가져오거나 설정합니다. (타입 인수 존재시 포함)
        /// 이 프로퍼티는 최상위 타입에 대해 null을 반환합니다.
        /// </summary>
        new IType DeclaringType { get; } // solves ambiguity between IType.DeclaringType and IEntity.DeclaringType

        
        /// <summary>
        /// 이 타입이 지정된 인터페이스 구성원을 구현하는 방법을 결정합니다.
        /// </summary>
        /// <returns>
        /// The method on this type that implements the interface member;
        /// or null if the type does not implement the interface.
        /// </returns>
        IMember GetInterfaceImplementation(IMember interfaceMember);

        /// <summary>
        /// 이 타입이 지정된 인터페이스 구성원을 구현하는 방법을 결정합니다.
        /// </summary>
        /// <returns>
        /// For each interface member, this method returns the class member 
        /// that implements the interface member.
        /// For interface members that are missing an implementation, the
        /// result collection will contain a null element.
        /// </returns>
        IList<IMember> GetInterfaceImplementation(IList<IMember> interfaceMembers);
    }
}
