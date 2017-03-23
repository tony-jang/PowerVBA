using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public enum CodeType
    {
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

    }
}
