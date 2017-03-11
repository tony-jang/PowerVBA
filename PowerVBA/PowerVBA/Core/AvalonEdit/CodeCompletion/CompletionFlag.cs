using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.CodeCompletion
{
    enum CompletionFlag
    {
        /// <summary>
        /// 아무런 플래그가 없습니다.
        /// </summary>
        None = 0,
        /// <summary>
        /// 특정한 상황이 없습니다. 일반적인 상황에서 사용됩니다.
        /// </summary>
        General = 1,
        /// <summary>
        /// 선언을 위한 플래그 입니다.
        /// </summary>
        Declarator = 2,
        /// <summary>
        /// 변수의 타입을 정합니다.
        /// </summary>
        As = 4,
        /// <summary>
        /// Class나 Function 또는 Sub 또는 Type이 포함되어 있는 플래그입니다.
        /// </summary>
        DeclareType = 8,
        /// <summary>
        /// 네임스페이스를 따라 가는 상황입니다. 예 (aaa.bbb.ccc.)
        /// </summary>
        InNameSpace = 16,
        /// <summary>
        /// 이름을 지정하거나 추가적으로 선언문을 사용할 수 있습니다.
        /// </summary>
        Naming = 32,
        /// <summary>
        /// 메소드에서 이름을 지정합니다
        /// </summary>
        MethodNaming = 64,
        /// <summary>
        /// 오직 이름만 설정할 수 있습니다 (Dim)
        /// </summary>
        OnlyNaming = 128,
        /// <summary>
        /// 프로시져 내부에서 사용 가능합니다.
        /// </summary>
        InProcedure = 256
    }
}
