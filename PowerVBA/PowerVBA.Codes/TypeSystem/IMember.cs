using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 메소드/필드/속성/이벤트
    /// </summary>
    public interface IUnresolvedMember : IUnresolvedEntity, IMemberReference
    {
        /// <summary>
        /// Gets the return type of this member.
        /// This property never returns null.
        /// </summary>
        ITypeReference ReturnType { get; }
        
        /// <summary>
        /// Member를 해석합니다.
        /// </summary>
        new IMember Resolve(ITypeResolveContext context);

        /// <summary>
        /// 해석된 Member를 생성합니다.
        /// </summary>
        /// <param name="context">
        /// 부모 형식 정의를 포함하는 언어별 컨텍스트.
        /// <see cref="IUnresolvedTypeDefinition.CreateResolveContext"/>
        /// </param>
        IMember CreateResolved(ITypeResolveContext context);
    }


    public interface IMemberReference : ISymbolReference
    {
        /// <summary>
        /// 멤버의 선언된 타입 참조를 가져옵니다.
        /// </summary>
        ITypeReference DeclaringTypeReference { get; }

        /// <summary>
        /// 멤버를 해석합니다.
        /// </summary>
        new IMember Resolve(ITypeResolveContext context);
    }

    /// <summary>
    /// 메소드/필드/속성/이벤트
    /// </summary>
    public interface IMember : IEntity
    {
        /// <summary>
        /// 이 멤버의 원래 멤버 정의를 가져옵니다.
        /// Returns <c>this</c> if this is not a specialized member.
        /// Specialized members are the result of overload resolution with type substitution.
        /// </summary>
        IMember MemberDefinition { get; }

        /// <summary>
        /// 이 멤버가 생성된 해석되지 않은 멤버 인스턴스를 가져옵니다.
        /// 확인되지 않은 멤버 인스턴스가 없는 특수 멤버에 대해 null을 반환 할 수 있습니다.
        /// </summary>
        /// <remarks>
        /// For specialized members, this property returns the unresolved member for the original member definition.
        /// For partial methods, this property returns the implementing partial method declaration, if one exists, and the
        /// defining partial method declaration otherwise.
        /// For the members used to represent the built-in C# operators like "operator +(int, int);", this property returns <c>null</c>.
        /// </remarks>
        IUnresolvedMember UnresolvedMember { get; }

        /// <summary>
        /// 이 멤버의 반환 타입을 가져옵니다.
        /// 이 속성은 null을 반환 하지 않습니다.
        /// </summary>
        IType ReturnType { get; }

        /// <summary>
        /// 이 멤버로 구현된 인터페이스 멤버를 가져옵니다. (암시 or 명시 모두)
        /// </summary>
        IList<IMember> ImplementedInterfaceMembers { get; }
        
        
        /// <summary>
        /// Creates a member reference that can be used to rediscover this member in another compilation.
        /// </summary>
        /// <remarks>
        /// If this member is specialized using open generic types, the resulting member reference will need to be looked up in an appropriate generic context.
        /// Otherwise, the main resolve context of a compilation is sufficient.
        /// </remarks>
        [Obsolete("ToReference 메소드를 대신 사용하세요.")]
        IMemberReference ToMemberReference();

        /// <summary>
        /// Creates a member reference that can be used to rediscover this member in another compilation.
        /// </summary>
        /// <remarks>
        /// If this member is specialized using open generic types, the resulting member reference will need to be looked up in an appropriate generic context.
        /// Otherwise, the main resolve context of a compilation is sufficient.
        /// </remarks>
        new IMemberReference ToReference();

        /// <summary>
        /// Gets the substitution belonging to this specialized member.
        /// Returns TypeParameterSubstitution.Identity for not specialized members.
        /// </summary>
        //TypeParameterSubstitution Substitution
        //{
        //    get;
        //}

        /// <summary>
        /// Specializes this member with the given substitution.
        /// If this member is already specialized, the new substitution is composed with the existing substition.
        /// </summary>
        //IMember Specialize(TypeParameterSubstitution substitution);
    }
}
