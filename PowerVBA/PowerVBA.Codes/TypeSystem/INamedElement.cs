using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface INamedElement
    {
        /// <summary>
        /// 반환 형식이 가리키는 클래스의 정규화 된 이름을 가져옵니다.
        /// </summary>
        string FullName { get; }

        /// <summary>
        /// 반환 값의 형태를 가리키는 클래스의 짧은 이름을 가져옵니다.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 요소의 전체 리플렉션 이름을 가져옵니다.
        /// </summary>
        string ReflectionName { get; }

        /// <summary>
        /// 이 독립체가 포함 된 네임 스페이스의 전체 이름을 가져옵니다.
        /// </summary>
        string Namespace { get; }
    }
}
