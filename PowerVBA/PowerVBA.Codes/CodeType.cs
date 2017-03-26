using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public enum CodeType
    {
        /// <summary>
        /// 알 수 없는 코드입니다.
        /// </summary>
        Unknown,


        #region [  엑세서  ]

        /// <summary>
        /// Public ~~
        /// </summary>
        PublicAccessor,
        /// <summary>
        /// Private ~~
        /// </summary>
        PrivateAccessor,

        #endregion

        #region [  블록  ]
        /// <summary>Type 선언 블록 전체입니다.</summary>
        TypeBlock,

        /// <summary>Sub 선언 블록 전체입니다.</summary>
        SubBlock,
        
        /// <summary>Function 선언 블록 전체입니다.</summary>
        FunctionBlock,

        /// <summary>If 선언 블록 전체입니다.</summary>
        IfBlock,

        /// <summary>Select Case 선언 블록 전체입니다.</summary>
        SelectCaseBlock,

        /// <summary>Enum 선언 블록 전체입니다.</summary>
        EnumBlock,

        /// <summary>Property 선언 블록 전체입니다.</summary>
        PropertyBlock,

        /// <summary>For 선언 블록 전체입니다.</summary>
        ForBlock,

        /// <summary>For Each 선언 블록 전체입니다.</summary>
        ForEachBlock,

        /// <summary>Do 선언 블록 전체입니다.</summary>
        DoBlock,

        /// <summary>Do While 선언 블록 전체입니다.</summary>
        DoWhileBlock,

        /// <summary>Do Until 선언 블록 전체입니다.</summary>
        DoUntilBlock,
        
        /// <summary>While 선언 블록 전체입니다.</summary>
        WhileBlock,

        #endregion

        #region [  선언  ]

        /// <summary>
        /// Type 선언 부분입니다.
        /// </summary>
        DeclareType,

        /// <summary>
        /// Sub 선언 부분입니다.
        /// </summary>
        DeclareSub,

        /// <summary>
        /// Function 선언 부분입니다.
        /// </summary>
        DeclareFunction,
        

        /// <summary>
        /// If 선언 부분입니다.
        /// </summary>
        DeclareIf,

        /// <summary>
        /// Select Case 선언 부분입니다.
        /// </summary>
        DeclareSelectCase,

        /// <summary>
        /// Enum 선언 부분입니다.
        /// </summary>
        DeclareEnum,

        /// <summary>
        /// Property 선언 부분입니다.
        /// </summary>
        DeclareProperty,

        /// <summary>
        /// For 선언 부분입니다.
        /// </summary>
        DeclareFor,

        /// <summary>
        /// For Each 선언 부분입니다.
        /// </summary>
        DeclareForEach,

        /// <summary>
        /// Do 선언 부분입니다.
        /// </summary>
        DeclareDo,

        /// <summary>
        /// Do While 선언 부분입니다.
        /// </summary>
        DeclareDoWhile,

        /// <summary>
        /// Do Until 선언 부분입니다.
        /// </summary>
        DeclareDoUntil,

        /// <summary>
        /// While 선언 부분입니다.
        /// </summary>
        DeclareWhile,

        #endregion
        
        #region [  Close  ]

        /// <summary>
        /// End Sub
        /// </summary>
        EndSub,
        /// <summary>
        /// End Function
        /// </summary>
        EndFunction,
        /// <summary>
        /// End Enum
        /// </summary>
        EndEnum,
        /// <summary>
        /// End Property
        /// </summary>
        EndProperty,
        /// <summary>
        /// End If
        /// </summary>
        EndIf,
        /// <summary>
        /// Next
        /// </summary>
        Next,
        /// <summary>
        /// While
        /// </summary>
        While,
        /// <summary>
        /// Wend
        /// </summary>
        Wend,

        #endregion

        #region [  String  ]

        /// <summary>
        /// 문자열 내부입니다.
        /// </summary>
        String,
        /// <summary>
        /// 식별자를 나타냅니다.
        /// </summary>
        Identifier,

        #endregion

        /// <summary>
        /// 클래스를 나타냅니다. 해당 선언은 잘못되었습니다.
        /// </summary>
        Class,
        /// <summary>
        /// 모듈을 나타냅니다. 해당 선언은 잘못되었습니다.
        /// </summary>
        Module,
        Dim,
        /// <summary>
        /// - + / * 와 같은 연산자를 나타냅니다.
        /// </summary>
        Operator,
        ExitDo,
        
    }
}
