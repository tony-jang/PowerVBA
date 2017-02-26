using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Global.Regex
{
    /// <summary>
    /// 자주 사용하는 또는 긴 정규 표현식 변수들을 정의해둔 곳입니다.
    /// </summary>
    public static class Var
    {
        #region [  변수 (Variable)  ]

        /// <summary>
        /// 빈칸을 체크합니다. 즉, 빈칸을 체크 하는 변수입니다. 1개의 그룹을 가지고 있습니다.
        /// </summary>
        public static string blnkChk { get; } = @"(\t|\f|\v| )+";


        public static string blnkOnLast { get; } = @"(\t|\f|\v| |\b)+";

        /// <summary>
        /// 프로퍼티 또는 메소드의 이름을 체크합니다. 1개의 그룹을 가지고 있습니다.
        /// </summary>
        public static string name { get; } = @"([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)";

        #endregion
    }

    /// <summary>
    /// 자주 사용하는 정규 표현식들을 정의해둔 곳입니다.
    /// </summary>
    public static class Pattern
    {
        #region [  패턴 (Pattern)  ]

        /// <summary>
        /// 변수를 선언 하기 위한 패턴입니다.
        /// </summary>
        public static string VariableDeclarePattern { get; } = $@"(?i)(dim|private|public)(\t|\f|\v| |)+{Var.name}(\t|\f|\v| |)+(?i)as(\t|\f|\v| |)+(.+)(\t|\f|\v| |)+";



        /// <summary>
        /// g1:public/private g2:blank g3:type g4:blank g5:name g6:blank g7:parameter g8:parameter(withoutbracket)
        /// </summary>
        public static string lineStartPattern { get; } = $@"^(?i)(public|private){Var.blnkChk}(?i)(sub|type|function){Var.blnkChk}{Var.name}{Var.blnkOnLast
                                                     }(|\((|.+)\))$";




        /// <summary>
        /// 라인의 종료를 확인하기 위한 패턴입니다.
        /// </summary>
        // 1 : 확인 된 종료 선언 타입 (Sub or Type or Function)
        public static string lineEndPattern { get; } = $@"^(?i)end (sub|function|type|enum)$";



        /// <summary>
        /// 빈칸 (탭, 비어 있음, 한칸 띄우기 등) 을 확인하기 위한 패턴입니다.
        /// </summary>
        // 0 : Full Match
        public static string blankCheckPattern { get; } = $@"^(\t|\f|\v| |)$";

        #endregion

        #region 약한 체크 패턴 (Weak Check Pattern)

        /// <summary>
        /// g1:any g2:blank g3:any g4:blank g5:name g6:blank g7:parameter g8:parameter(Withoutbracket)
        /// </summary>
        public static string g_lineStartPattern { get; } = $@"^(.+){Var.blnkChk}(.+){Var.blnkChk}{Var.name}{Var.blnkOnLast}(|\((|.+)\))";

        public static string g_lineEndPattern { get; } = $@"(.+){Var.blnkChk}(.+)";

        #endregion
    }
}
