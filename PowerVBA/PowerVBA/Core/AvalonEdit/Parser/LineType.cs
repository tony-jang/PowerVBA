using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Parser
{
    enum LineType
    {
        /// <summary>
        /// 알 수 없는 줄입니다.
        /// </summary>
        Unknown,

        /// <summary>
        /// 주석
        /// </summary>
        Remark,
        /// <summary>
        /// 코드 정렬
        /// </summary>
        Region,

        /// <summary>
        /// 프로시져 선언
        /// </summary>
        ProcedureStart,
        /// <summary>
        /// 프로시져 선언 종료
        /// </summary>
        ProcedureEnd,


        /// <summary>
        /// Enum 선언 시작
        /// </summary>
        EnumStart,
        /// <summary>
        /// Enum 선언 종료
        /// </summary>
        EnumEnd,


        /// <summary>
        /// Type 선언 시작
        /// </summary>
        TypeStart,
        /// <summary>
        /// Type 선언 종료
        /// </summary>
        TypeEnd,


        /// <summary>
        /// 글로벌 변수 선언
        /// </summary>
        GlobalVariable,
        /// <summary>
        /// 지역 변수 선언
        /// </summary>
        LocalVariable,
        /// <summary>
        /// 이벤트 변수 선언
        /// </summary>
        EventVariable,


        /// <summary>
        /// 전처리기 지시문
        /// </summary>
        Preprocessor,
        /// <summary>
        /// 값 대입
        /// </summary>
        ValueAssign,


        /// <summary>
        /// 함수 호출
        /// </summary>
        CallFunction,


        /// <summary>
        /// If문 선언
        /// </summary>
        If,
        /// <summary>
        /// ElseIf문 선언
        /// </summary>
        ElseIf,
        /// <summary>
        /// End If문 선언
        /// </summary>
        EndIf,
        /// <summary>
        /// Else문 선언
        /// </summary>
        Else,


        /// <summary>
        /// Select Case문 선언
        /// </summary>
        SelectCase,
        /// <summary>
        /// Case 선언
        /// </summary>
        Case,
        /// <summary>
        /// Case Else 선언
        /// </summary>
        CaseElse,
        /// <summary>
        /// End Select 선언
        /// </summary>
        EndSelect,


        /// <summary>
        /// Do 선언
        /// </summary>
        Do,
        /// <summary>
        /// Do While 선언
        /// </summary>
        DoWhile,
        /// <summary>
        /// Do Until 선언
        /// </summary>
        DoUntil,


        /// <summary>
        /// Option Explicit 설정
        /// </summary>
        OptionExplicit,
        /// <summary>
        /// Option Compare 설정
        /// </summary>
        OptionCompare,
        /// <summary>
        /// Option Base 설정
        /// </summary>
        OptionBase,
        /// <summary>
        /// Option Private 설정
        /// </summary>
        OptionPrivate,


        /// <summary>
        /// 다중 라인
        /// </summary>
        MultiLine

    }
}
