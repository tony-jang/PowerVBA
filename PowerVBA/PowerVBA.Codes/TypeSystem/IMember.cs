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
        /// Gets the unresolved member instance from which this member was created.
        /// This property may return <c>null</c> for special members that do not have a corresponding unresolved member instance.
        /// </summary>
        /// <remarks>
        /// For specialized members, this property returns the unresolved member for the original member definition.
        /// For partial methods, this property returns the implementing partial method declaration, if one exists, and the
        /// defining partial method declaration otherwise.
        /// For the members used to represent the built-in C# operators like "operator +(int, int);", this property returns <c>null</c>.
        /// </remarks>
        IUnresolvedMember UnresolvedMember { get; }

        /// <summary>
        /// Gets the return type of this member.
        /// This property never returns <c>null</c>.
        /// </summary>
        IType ReturnType { get; }

        /// <summary>
        /// Gets the interface members implemented by this member (both implicitly and explicitly).
        /// </summary>
        IList<IMember> ImplementedInterfaceMembers { get; }

        /// <summary>
        /// Gets whether this member is explicitly implementing an interface.
        /// </summary>
        bool IsExplicitInterfaceImplementation { get; }

        /// <summary>
        /// Gets if the member is virtual. Is true only if the "virtual" modifier was used, but non-virtual
        /// members can be overridden, too; if they are abstract or overriding a method.
        /// </summary>
        bool IsVirtual { get; }

        /// <summary>
        /// Gets whether this member is overriding another member.
        /// </summary>
        bool IsOverride { get; }

        /// <summary>
        /// Gets if the member can be overridden. Returns true when the member is "abstract", "virtual" or "override" but not "sealed".
        /// </summary>
        bool IsOverridable { get; }

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
        TypeParameterSubstitution Substitution
        {
            get;
        }

        /// <summary>
        /// Specializes this member with the given substitution.
        /// If this member is already specialized, the new substitution is composed with the existing substition.
        /// </summary>
        IMember Specialize(TypeParameterSubstitution substitution);
    }
}
