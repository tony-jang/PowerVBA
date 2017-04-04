using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    enum Keywords
    {
        /// <summary>
        /// 변수 선언 뒤의 타입을 나타냅니다.
        /// 예 : Dim A As String
        /// </summary>
        As,
        /// <summary>
        /// 프로그램을 종료하거나 블록을 종료합니다.
        /// 예. End / End Sub
        /// </summary>
        End,
        /// <summary>
        /// 현재 위치에서의 문법을 나갑니다.
        /// </summary>
        Exit,

        /// <summary>
        /// If문, Case Else 등으로 사용 할 수 있습니다.
        /// </summary>
        Else,

        /// <summary>
        /// For Each문, For문 등으로 사용 할 수 있습니다.
        /// </summary>
        For,
        /// <summary>
        /// For문 뒤에만 올 수 있는 Each입니다.
        /// </summary>
        Each,
        /// <summary>
        /// If문 / End If문으로 사용 할 수 있습니다.
        /// </summary>
        If,
        /// <summary>
        /// Select문 / End Select로 사용 할 수 있습니다.
        /// </summary>
        Select,
        /// <summary>
        /// Case &lt;Expression>입니다.
        /// </summary>
        Case,
        Do,
        While,
        Until,
        Wend,

        Sub,
        Function,

        Enum,

    }
}
