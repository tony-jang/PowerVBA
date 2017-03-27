using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{

    public enum Accessibility : byte
    {

        /// <summary>
        /// 이 독립체는 완전히 접근이 불가능합니다.
        /// </summary>
        None,
        /// <summary>
        /// 이 독립체는 같은 클래스에서만 접근 가능합니다.
        /// </summary>
        Private,
        /// <summary>
        /// 이 독립체는 어디든 엑세스 할 수 있습니다.
        /// </summary>
        Public
    }

    /// <summary>
    /// 엑세서를 확인하는 인터페이스입니다.
    /// </summary>
    public interface IHasAccessibility
    {
        Accessibility Accessibility { get; }

        bool IsPrivate { get; }

        bool IsPublic { get; }
    }
}
