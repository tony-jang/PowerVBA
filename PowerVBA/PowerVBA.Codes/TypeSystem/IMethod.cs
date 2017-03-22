using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IUnresolvedMethod : IUnresolvedParameterizedMember
    {

        IList<IUnresolvedTypeParameter> TypeParameters { get; }

        bool IsConstructor { get; }
        bool IsDestructor { get; }
        bool IsOperator { get; }
        

        [Obsolete("Use IsPartial && !HasBody instead")]
        bool IsPartialMethodDeclaration { get; }

        [Obsolete("Use IsPartial && HasBody instead")]
        bool IsPartialMethodImplementation { get; }
        

        /// <summary>
        /// 이 메소드가 접근자(엑세서)라면 해당 속성/이벤트에 대한 참조를 반환합니다.
        /// 그 이외엔 null을 반환합니다.
        /// </summary>
        IUnresolvedMember AccessorOwner { get; }

        /// <summary>
        /// member를 해석합니다.
        /// </summary>
        new IMethod Resolve(ITypeResolveContext context);
    }

    /// <summary>
    /// 메소드, 생성자, 파괴자, 연산자를 나타냅니다.
    /// </summary>
    public interface IMethod : IParameterizedMember
    {
        /// <summary>
        /// 해석되지 않은 메소드 부분을 가져옵니다.
        /// </summary>
        IList<IUnresolvedMethod> Parts { get; }

        /// <summary>
        /// 이 메소드의 형태 파라미터를 가져옵니다. 메소드가 일반적이지 않은 경우는 빈 상태의 리스트가 됩니다.
        /// </summary>
        IList<ITypeParameter> TypeParameters { get; }

        /// <summary>
        /// 이 메소드가 매개 변수화 된 제네릭 메소드인지 여부를 가져옵니다.
        /// </summary>
        bool IsParameterized { get; }

        /// <summary>
        /// 이 메소드에 전달 된 형식 인수를 가져옵니다.
        /// 메소드가 Generic이나 아직 매개 변수화 되지 않은 경우에 메소드는 메소드가 자체 유형 인수로 매개 변수화 된 것 처럼 유형 매개 변수를 리턴합니다.
        /// 
        /// NOTE: The type will change to IReadOnlyList&lt;IType&gt; in future versions.
        /// </summary>
        IList<IType> TypeArguments { get; }
        
        bool IsConstructor { get; }
        bool IsDestructor { get; }
        bool IsOperator { get; }



        /// <summary>
        /// 메서드가 속성 / 이벤트 접근 자인지 여부를 가져옵니다.
        /// </summary>
        bool IsAccessor { get; }

        /// <summary>
        /// 이 메소드 접근자라면 해당 속성/이벤트를 반환합니다.
        /// Otherwise, returns null.
        /// </summary>
        IMember AccessorOwner { get; }

        /// <summary>
        /// If this method is reduced from an extension method return the original method, <c>null</c> otherwise.
        /// A reduced method doesn't contain the extension method parameter. That means that has one parameter less than it's definition.
        /// </summary>
        //IMethod ReducedFrom { get; }

        /// <summary>
        /// Specializes this method with the given substitution.
        /// If this method is already specialized, the new substitution is composed with the existing substition.
        /// </summary>
        new IMethod Specialize(TypeParameterSubstitution substitution);
    }
}
