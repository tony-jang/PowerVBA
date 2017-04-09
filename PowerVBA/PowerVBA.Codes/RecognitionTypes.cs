using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    enum RecognitionTypes
    {
        /// <summary>
        /// 끝이라는 기준 상관 없이 인식 단 String 내부가 아닌 곳에서 ':'가 인식되면 중단
        /// </summary>
        ToEndLine,
        /// <summary>
        /// 파라미터에서의 식 이 경우 ')'나 ','에 반응하여 중단
        /// </summary>
        Parameter,

        /// <summary>
        /// If나 ElseIf에서 Then 이후 절까지 인식 이 경우 식이 없는데 Then이 나오면 오류로 인식
        /// </summary>
        BeforeThen,

        /// <summary>
        /// 괄호 여는 식 인식.
        /// 예) (A = B) + (B = C)의 경우 (A = B) 까지만 인식하여 반응
        /// </summary>
        BracketExpression,
    }
}
