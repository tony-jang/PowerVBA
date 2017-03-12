using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PowerVBA.Global.RegexExpressions
{
    static class CodePattern
    {
        public class RegexString
        {
            public RegexString(string str)
            {
                String = str;
            }
            public string String { get; set; }

            public static implicit operator RegexString(string v)
            {
                return new RegexString($"(?i)^{BlankNull}{v}$");
            }

            public static implicit operator string(RegexString regStr)
            {
                return regStr.String;
            }


            public bool IsMatch(string s)
            {
                return Regex.IsMatch(s, String);
            }

            public Match Match(string s)
            {
                return Regex.Match(s, String);
            }
        }

        #region [  변수  ]

        public static string Blank = @"\s+";

        public static string BlankNull = @"\s*";

        /// <summary>
        /// 사용 가능한 이름입니다. (1 Group)
        /// </summary>
        public static string Name = @"([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)";

        /// <summary>
        /// - 범위의 숫자와 + 범위의 숫자입니다. (1 Group)
        /// </summary>
        public static string Digit = @"(-\d+|\d+)";

        public static string GetNegativeLookBehind(string[] strings)
        {
            return $"(?<!{string.Join("|", strings)})";
        }
        public static string GetNegativeLookBehind(string String)
        {
            return $"(?<!{String})";
        }


        #endregion

        #region [  기본 패턴  ]

        // 아무것도 넣지 않아도 자동 암묵적 변환에서 원하는 패턴이 만들어짐
        public static RegexString BlankNullPattern { get; } = "";

        public static RegexString Default { get; } = "[a-zA-Z_]";

        public static RegexString Second { get; } = $"{Name}{Blank}";

        #endregion


        #region [  For Standard  ]

        /// <summary>
        /// [For] To 이전 절까지를 인식합니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Digit
        public static RegexString ForBeforeTo { get; } = $@"For{Blank}{Name}{Blank}={Blank}{Digit}{Blank}";


        /// <summary>
        /// [For] Step 이전 절까지를 인식합니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Digit (StartPos) Group 3 : Digit (EndPos)
        public static RegexString ForBeforeStep { get; } = $@"For{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{Blank}";


        /// <summary>
        /// [For][완] For i = 1 to 10 과 같은 문법을 인식합니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Digit (StartPos) Group 3 : Digit (EndPos)
        public static RegexString ForStandard { get; } = $@"For{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{BlankNull}";

        /// <summary>
        /// [For][완] Step까지 작성한 기본 For문 완성입니다. Ex) For i = 1 to 10 Step 2
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Digit (StartPos) Group 3 : Digit (EndPos) Group 4 : Digit (Step)
        public static RegexString ForStep { get; } = $@"For{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{Blank}Step{Blank}{Digit}{BlankNull}";


        /// <summary>
        /// [For][완] Next (For문 닫기) 부분입니다.
        /// </summary>
        public static RegexString ForNext { get; } = $@"Next{BlankNull}";
        #endregion

        #region [  For Extension  ]


        /// <summary>
        /// [For Ex] As 전까지 인식시킵니다.
        /// </summary>
        // Group 1 : Name(Ref Var)
        public static RegexString ForBeforeAs_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}";


        /// <summary>
        /// [For Ex] Type 전까지 인식시킵니다.
        /// </summary>
        // Group 1 : Name(Ref Var)
        public static RegexString ForBeforeType_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}As{Blank}";

        /// <summary>
        /// [For Ex] To 전까지 인식시킵니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Name(Type) Group 3 : Digit (StartPos)
        public static RegexString ForBeforeTo_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}As{Blank}{Name}{Blank}={Blank}{Digit}{Blank}";

        /// <summary>
        /// [For Ex] Step 전까지 인식시킵니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Name(Type) Group 3 : Digit (StartPos) Group 4 : Digit (EndPos)
        public static RegexString ForBeforeStep_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}As{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{Blank}";


        /// <summary>
        /// [For Ex][완] For i As Integer = 1 to 10 과 같은 문법을 인식합니다.
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Name(Type) Group 3 : Digit (StartPos) Group 4 : Digit (EndPos)
        // To VBA Grammer :
        // Dim i As Integer
        // For i = 1 to 10
        public static RegexString ForStandard_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}As{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{BlankNull}";


        /// <summary>
        /// [For Ex][완] Step까지 작성한 기본 For문 완성입니다. Ex) For i As Integer = 1 to 10 Step 2
        /// </summary>
        // Group 1 : Name(Ref Var) Group 2 : Name(Type) Group 3 : Digit (StartPos)
        // To VBA Grammer :
        // Dim i As Integer
        // For i = 1 to 10 Step 2
        public static RegexString ForStep_Ex { get; } = $@"{BlankNull}For{Blank}{Name}{GetNegativeLookBehind("Each")}{Blank}As{Blank}{Name}{Blank}={Blank}{Digit}{Blank}To{Blank}{Digit}{Blank}Step{Blank}{Digit}{BlankNull}";


        #endregion

        #region [  For Each  ]

        public static RegexString ForEach_BeforeEach { get; } = $@"For{Blank}";

        #endregion


        #region [  Select Case  ]

        /// <summary>
        /// [Select Case] Case 전을 인식합니다.
        /// </summary>
        public static RegexString SCBeforeCase { get; } = $@"Select{Blank}";


        /// <summary>
        /// [Select Case] Object 전까지를 인식합니다.
        /// </summary>
        public static RegexString SCBeforeObject { get; } = $@"Select{Blank}Case{Blank}";


        /// <summary>
        /// [Select Case][완] Select Case 기본 문을 인식합니다.
        /// </summary>
        public static RegexString SCStandard { get; } = $@"Select{Blank}Case{Blank}(.+){BlankNull}";


        /// <summary>
        /// [Select Case] Case Object 전까지를 인식합니다.
        /// </summary>
        public static RegexString SCBeforeCaseObject { get; } = $@"Case{Blank}";


        /// <summary>
        /// [Select Case][완] Case 기본 문을 인식합니다.
        /// </summary>
        public static RegexString SCCase { get; } = $@"Case{Blank}(.+){BlankNull}";


        /// <summary>
        /// [Select Case][완] Case Else문을 인식합니다.
        /// </summary>
        public static RegexString SCCaseElse { get; } = $@"Case{Blank}Else{BlankNull}";


        #endregion



        #region [  Do  ]

        /// <summary>
        /// [Do][완] Do문을 인식합니다.
        /// </summary>
        public static RegexString DoStandard { get; } = $"Do{BlankNull}";

        /// <summary>
        /// [Do][완] Do문의 Loop을 인식합니다.
        /// </summary>
        public static RegexString LoopStandard { get; } = $"Loop{BlankNull}";

        /// <summary>
        /// [Do] Loop에서 Until이나 While 인식 전까지를 인식합니다.
        /// </summary>
        public static RegexString LoopBeforeWhileUntil { get; } = $"Loop{Blank}";

        /// <summary>
        /// [Do] Do에서 Until이나 While 인식 전까지를 인식합니다.
        /// </summary>
        public static RegexString DoBeforeWhileUntil { get; } = $"Do{Blank}";

        #endregion

        #region [  Do While & Loop While  ]

        /// <summary>
        /// [Do While] Do While에서 식 전까지를 인식합니다.
        /// </summary>
        public static RegexString DoWhileBeforeObject { get; } = $"Do{Blank}While{Blank}";

        /// <summary>
        /// [Do While][완] Do While문을 인식합니다.
        /// </summary>
        public static RegexString DoWhileStandard { get; } = $"Do{Blank}While{Blank}(.+){BlankNull}";

        /// <summary>
        /// [Do While] Loop While 문의 식 전까지를 인식합니다.
        /// </summary>
        public static RegexString LoopWhileBeforeExp { get; } = $"Loop{Blank}While{Blank}";


        /// <summary>
        /// [Do While][완] Loop While 문을 인식합니다.
        /// </summary>
        public static RegexString LoopWhile { get; } = $"Loop{Blank}While{Blank}(.+){BlankNull}";

        #endregion

        #region [  Do Until & Loop Until  ]

        /// <summary>
        /// [Do Until] Do Until에서 식 전까지를 인식합니다.
        /// </summary>
        public static RegexString DoUntilBeforeObject { get; } = $"Do{Blank}Until{Blank}";

        /// <summary>
        /// [Do Until][완] Do Until문을 인식합니다.
        /// </summary>
        public static RegexString DoUntilStandard { get; } = $"Do{Blank}Until{Blank}(.+){BlankNull}";

        /// <summary>
        /// [Do Until] Loop Until 문의 식 전까지를 인식합니다.
        /// </summary>
        public static RegexString LoopUntilBeforeExp { get; } = $"Loop{Blank}Until{Blank}";

        /// <summary>
        /// [Do Until][완] Loop Until 문을 인식합니다.
        /// </summary>
        public static RegexString LoopUntil { get; } = $"Loop{Blank}Until{Blank}(.+){BlankNull}";

        #endregion

        #region [  While  ]

        public static RegexString WhileBeforeExp { get; } = $"While{Blank}";

        /// <summary>
        /// [While Ex] End While문을 인식합니다. (단, 보조 VBA 파일 사용 프로젝트 일시에만 가능)
        /// </summary>
        // To VBA Grammer : Wend
        public static RegexString EndWhile_Ex { get; } = $"End While{BlankNull}";

        /// <summary>
        /// [While] Wend문을 인식합니다.
        /// </summary>
        public static RegexString Wend { get; } = $"Wend{BlankNull}";

        #endregion



        #region [  For, Do, While, Select Case  ]

        /// <summary>
        /// [For/Do/While/Select Case] Exit For, Do, While 이전 까지를 인식합니다.
        /// </summary>
        public static RegexString ExitGrammers { get; } = $"Exit{Blank}";

        /// <summary>
        /// [For/Do/While] Continue For, Do, While 이전 까지를 인식합니다.
        /// </summary>
        public static RegexString ContinueGrammers { get; } = $"Continue{Blank}";

        #endregion


        #region [  Variable  ]

        public static RegexString VarBeforeNaming { get; } = $@"Dim{Blank}";

        public static RegexString VarNaming { get; } = $@"Dim{Blank}{Name}";

        public static RegexString VarBefore_As { get; } = $@"Dim{Blank}{Name}{Blank}";

        public static RegexString VarBefore_Type { get; } = $@"Dim{Blank}{Name}{Blank}As{Blank}";

        #endregion


        #region [  Method  ]



        #endregion

    }
}
