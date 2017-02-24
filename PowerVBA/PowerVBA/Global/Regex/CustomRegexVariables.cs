using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Global.Regex
{
    /// <summary>
    /// 자주 사용하는 또는 긴 정규 표현식들을 정의해둔 곳입니다.
    /// </summary>
    public static class Variable
    {
        #region [  변수 (Variable)  ]

        /// <summary>
        /// 빈칸을 체크합니다. 즉, 빈칸을 체크 하는 변수입니다.
        /// </summary>
        public static string blnkChkVar { get; } = @"(\t|\f|\v| |)+";

        /// <summary>
        /// 프로퍼티 또는 메소드의 이름을 체크합니다.
        /// </summary>
        public static string nameVar { get; } = @"([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)";

        #endregion
    }

    public static class Pattern
    {
        #region [  패턴 (Pattern)  ]

        /// <summary>
        /// 변수를 선언 하기 위한 패턴입니다.
        /// </summary>
        public static string VariableDeclarePattern { get; } = @"(?i)(dim|private|public)(\t|\f|\v| |)+([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)(\t|\f|\v| |)+(?i)as(\t|\f|\v| |)+(.+)(\t|\f|\v| |)+";



        /// <summary>
        /// 라인의 시작, 함수등을 선언하기 위한 패턴입니다.
        /// </summary>
        public static string lineStartPattern { get; } = $@"^(?i)(public|private){Variable.blnkChkVar
                                                     }(?i)(sub|type|function){Variable.blnkChkVar}{Variable.nameVar}{Variable.blnkChkVar
                                                     }(|\((|.+)\))$";


        /// <summary>
        /// 라인의 종료를 확인하기 위한 패턴입니다.
        /// </summary>
        public static string lineEndPattern { get; } = $@"^(?i)end (sub|function|type|enum)$";



        /// <summary>
        /// 빈칸 (탭, 비어 있음, 한칸 띄우기 등) 을 확인하기 위한 패턴입니다.
        /// </summary>
        public static string blankCheckPattern { get; } = $@"^{Variable.blnkChkVar}$";

        #endregion
    }
}
