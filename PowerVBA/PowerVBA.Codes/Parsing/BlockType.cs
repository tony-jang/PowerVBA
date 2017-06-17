using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Parsing
{
    public enum BlockType
    {
        /// <summary>
        /// 알수 없는 블록을 의미합니다.
        /// </summary>
        Unknown,

        /// <summary>
        /// 모든 것을 포함하고 있는 전체 블록을 의미합니다.
        /// </summary>
        Both,

        /// <summary>
        /// Enum 블록입니다.
        /// </summary>
        Enum,
        /// <summary>
        /// 부분별로 나눈 블록입니다. #region과 #endregion으로 나눕니다.
        /// </summary>
        Region,
        /// <summary>
        /// 함수 블록입니다.
        /// </summary>
        Function,
        /// <summary>
        /// 서브 블록입니다.
        /// </summary>
        Sub,
        
        /// <summary>
        /// If 블록입니다. If와 End If 를 기록합니다.
        /// </summary>
        If,

        /// <summary>
        /// Select 블록입니다. Select와 End Select를 기준으로 나눕니다.
        /// </summary>
        Select,
        /// <summary>
        /// Case/Case Else 블록입니다. Case 블록의 끝나는 줄은 끝 부분에 Length 0으로 존재합니다.
        /// </summary>
        Case,
        /// <summary>
        /// Do 루프 블록입니다. Do While/Do Until 모두 이 블록을 사용합니다.
        /// </summary>
        Do,
        /// <summary>
        /// While 루프 블록입니다.
        /// </summary>
        While,
        

    }
}
