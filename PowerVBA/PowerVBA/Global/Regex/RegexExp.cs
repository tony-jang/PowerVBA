using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.RegexPattern
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

        /// <summary>
        /// 빈칸을 마지막에 체크합니다. Word Boundary가 포함되어 있습니다.
        /// </summary>
        public static string blnkOnLast { get; } = @"(\t|\f|\v| |\b)+";

        /// <summary>
        /// 빈칸이거나 아예 없을수도 있습니다.
        /// </summary>
        public static string blnkOrNull { get; } = @"(\t|\f|\v| |)+";

        /// <summary>
        /// 프로퍼티 또는 메소드의 이름을 체크합니다. 1개의 그룹을 가지고 있습니다.
        /// </summary>
        public static string name { get; } = @"([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)";

        /// <summary>
        /// a-zA-Z의 알파벳 문자를 가지고 있습니다.
        /// </summary>
        public static string Alphabet { get; } = "[a-zA-Z]";

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

        #region [  약한 체크 패턴 (Weak Check Pattern)  ]

        /// <summary>
        /// g1:any g2:blank g3:any g4:blank g5:name g6:blank g7:parameter g8:parameter(Withoutbracket)
        /// </summary>
        public static string g_lineStartPattern { get; } = $@"^(.+){Var.blnkChk}(.+){Var.blnkChk}{Var.name}{Var.blnkOnLast}(|\((|.+)\))";

        /// <summary>
        /// g1:any g2:blank g3:any
        /// </summary>
        public static string g_lineEndPattern { get; } = $@"(.+){Var.blnkChk}(.+)";

        #endregion


        #region [  띄울수 있는 패턴  ]

        // 특화된 것부터 배치


        /// <summary>
        /// [(dim|with)][blank]
        /// </summary>
        public static string pattern1 { get; } = $@"(?i)^{Var.blnkOrNull}(dim|with){Var.blnkChk}$";
        /// <summary>
        /// [(public|private)][blank]
        /// </summary>
        public static string pattern1_1 { get; } = $@"(?i)^{Var.blnkOrNull}(public|private){Var.blnkChk}$";
        /// <summary>
        /// [blank][alphabet+]
        /// </summary>
        public static string pattern1_2 { get; } = $@"(?i)^{Var.blnkOrNull}({Var.Alphabet}+)$";
        /// <summary>
        /// [dim|public|private][Tab][Name][function|sub|type|enum][blank]
        /// </summary>
        public static string pattern1_3 { get; } = $@"(?i)^(\t|\f|\v| |)+(dim|public|private)(\t|\f|\v| )+{Var.name}(?<!function|sub|type|enum){Var.blnkChk}$";
        /// <summary>
        /// [blank]
        /// </summary>
        public static string pattern2 { get; } = $@"^{Var.blnkOrNull}$";

        /// <summary>
        /// [private/public/null][blank][class/enum/type][blank]
        /// </summary>
        public static string pattern3 { get; } = $@"(?i)^{Var.blnkOrNull}(public|private|){Var.blnkOrNull}(class|enum|type){Var.blnkChk}$";
        /// <summary>
        /// [tab][dim|public|private][tab][name][function|sub|type|enum][blank]As[blank]
        /// </summary>
        public static string pattern4 { get; } = $@"(?i)(\t|\f|\v| |)+(dim|public|private)(\t|\f|\v| )+{Var.name}(?<!function|sub|type|enum){Var.blnkChk}As{Var.blnkChk}";
        #endregion
    }
}
