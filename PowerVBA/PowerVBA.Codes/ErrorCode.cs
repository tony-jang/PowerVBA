using PowerVBA.Codes.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public enum ErrorCode
    {
        /// <summary>커스텀 오류</summary>
        [CanReplace(1)]
        [KoError("%1")]
        VB0000 = 0,

        #region [  지원하지 않는 문법  ]

        /// <summary>미지원 문법 - Class</summary>
        [NotSupported]
        [KoError("'Class'는 VBA에서 지원하지 않습니다. 대신 Type을 이용해보세요.")]
        VB0001 = 1,

        /// <summary>미지원 문법 - Module</summary>
        [NotSupported]
        [KoError("'Module'은 VBA에서 지원하지 않습니다. 대신 모듈 파일을 생성해보세요.")]
        VB0002,

        /// <summary>미지원 문법 - Get</summary>
        [NotSupported]
        [KoError("'Get'은 VBA에서 지원하지 않습니다. 대신 Public Property Get을 이용해보세요.")]
        VB0003,
        
        /// <summary>미지원 문법 - End While</summary>
        [NotSupported]
        [KoError("'End While'은 VBA에서 지원하지 않습니다. 대신 Wend를 이용해보세요.")]
        VB0005,
        /// <summary>미지원 문법 - ReadOnly</summary>
        [NotSupported]
        [KoError("'ReadOnly'는 VBA에서 지원하지 않습니다.")]
        VB0006,
        /// <summary>미지원 문법 - AddHandler</summary>
        [NotSupported]
        [KoError("'AddHandler'는 VBA에서 지원하지 않습니다.")]
        VB0007,
        /// <summary>미지원 문법 - 선언과 동시에 초기화 불가</summary>
        [NotSupported]
        [KoError("VBA에서는 변수 선언과 동시에 초기화를 지원하지 않습니다.")]
        VB0008,
        /// <summary>미지원 문법 - 연산자 미지원</summary>
        [NotSupported]
        [CanReplace(2)]
        [KoError("VBA에서는 %1와(과) 같은 연산자를 지원하지 않습니다. 대신 %2로 사용 해보세요.")]
        VB0009,
        
        [CanReplace(1)]
        [KoError("Exit 뒤에는 '%1' 키워드 대신 Do, For, Sub, Function 또는 Property가 와야 합니다.")]
        VB0011,

        [NotSupported]
        [KoError("Return으로 값을 반환하는 것은 VBA에서 지원하지 않습니다. Identifier = Value를 사용하세요.")]
        VB0012,
        #endregion

        #region [  중복  ]

        /// <summary>선언문 중복</summary>
        [KoError("이미 선언 문이 작성되었습니다.")]
        VB0021 = 21,
        /// <summary>지정자 중복</summary>
        [KoError("지정자가 중복되었습니다.")]
        VB0022,

        #endregion


        #region [  위치적 사용 불가  ]

        /// <summary>사용 불가</summary>
        [CanReplace(3)]
        [KoError("%1은(는) %2와(과) 같은 %3가 올 수 없습니다.")]
        VB0025 = 25,

        /// <summary>사용 필요한 것 ~~뒤(1) ~~와 같은 (2) ~~가 와야함(3)</summary>
        [CanReplace(3)]
        [KoError("%1뒤에는 %2와(과) 같은 %3가 와야합니다.")]
        VB0026,
        
        /// <summary>식별자 뒤 특정 단어 와야 됨 (1)</summary>
        [CanReplace(1)]
        [KoError("식별자 뒤에는 %1가 와야 합니다.")]
        VB0027,

        /// <summary>특수 절 (1) 사용 불가</summary>
        [CanReplace(1)]
        [KoError("%1 절은 맨 처음에만 올 수 있습니다.")]
        VB0040 = 40,
        #endregion


        #region [  특수문자 사용 불가  ]

        /// <summary>특수문자 사용 불가</summary>
        [CanReplace(1)]
        [KoError("해당 특수문자 '%1'은(는) 사용할 수 없습니다.")]
        VB0041,

        /// <summary>전처리기 지시문 특수문자 사용 불가</summary>
        [KoError("전처리기 지시문 특수문자 '#'은(는) 사용할 수 없습니다.")]
        VB0042,


        /// <summary>해당 문자가 잘못되었습니다.</summary>
        [KoError("해당 문자가 잘못되었습니다.")]
        VB0043,

        /// <summary>해당 문자가 잘못되었습니다.</summary>
        [KoError("'+'나 '&'는 문자나 숫자를 이어줄 때만 사용 할 수 있습니다.")]
        VB0044,

        #endregion

        #region [  예약어 사용 불가  ]

        /// <summary>식별자 뒤 예약어 사용불가능 (1)</summary>
        [CanReplace(1)]
        [KoError("식별자 뒤에는 '%1'와 같은 예약어가 올 수 없습니다.")]
        VB0045 = 45,
        /// <summary>배열 크기 뒤에는 예약어 사용 불가</summary>
        [KoError("배열 크기 선언에는 예약어가 올 수 없습니다.")]
        VB0046,

        /// <summary>라벨 선언에는 예약어 사용 불가</summary>
        [CanReplace(1)]
        [KoError("라벨 선언에는 '%1'와 같은 예약어가 올 수 없습니다.")]
        VB0047,

        /// <summary>현재 컨텍스트에서 예약어 유효하지 않음</summary>
        [CanReplace(1)]
        [KoError("현재 컨텍스트에서는 '%1'와 같은 예약어가 유효하지 않습니다.")]
        VB0048,


        #endregion


        #region [  괄호 사용 불가  ]

        /// <summary>여는 괄호 현위치 사용 불가</summary>
        [KoError("여는 괄호는 현재 위치에서 사용 할 수 없습니다.")]
        VB0050 = 50,
        
        /// <summary>여는 괄호 필요</summary>
        [KoError(") 를 사용하기 위해서는 (가 필요합니다.")]
        VB0051,

        /// <summary>닫는 괄호 누락 갯수 (1)</summary>
        [CanReplace(1)]
        [KoError("')' %1개가 누락되었습니다.")]
        VB0052,

        #endregion


        #region [  End 오류  ]

        /// <summary>End로 닫을 수 없는 아이템 (1)</summary>
        [CanReplace(1)]
        [KoError("%1은(는) End로 닫을 수 없습니다. If, Select, Sub, Function, Property, Type, With, Enum 또는 그냥 End로만 사용해야 합니다.")]
        VB0055 = 55,

        /// <summary>처음에 End가 나오지 않은 오류</summary>
        [KoError("End는 맨 처음에만 올 수 있습니다.")]
        VB0056,

        [KoError("닫는 키워드는 End 바로 다음에만 올 수 있습니다.")]
        VB0057,
        #endregion

        #region [  연산자 오류  ]

        /// <summary>저장시 자동 치환 경고</summary>
        [CanReplace(2)]
        [KoError("%1은(는) 저장할때 %2로 자동으로 바뀝니다.")]
        VB0060 = 60,

        /// <summary>연산자 오류</summary>
        [CanReplace(1)]
        [KoError("%1은(는) 올바른 연산자가 아닙니다.")]
        VB0061,

        /// <summary>피연산자 오류</summary>
        [CanReplace(1)]
        [KoError("'%1' 키워드는 피연산자 다음에는 나올 수 없습니다.")]
        VB0062,
        /// <summary>연산자 오류</summary>
        [CanReplace(1)]
        [KoError("'%1' 키워드는 연산자 다음에는 나올 수 없습니다.")]
        VB0063,
        /// <summary>연산자 오류</summary>
        [CanReplace(1)]
        [KoError("'%1' 식은 연산자 다음에는 나올 수 없습니다.")]
        VB0064,
        /// <summary>연산자 오류</summary>
        [CanReplace(1)]
        [KoError("'%1' 식은 피연산자 다음에는 나올 수 없습니다.")]
        VB0065,
        [CanReplace(1)]
        [KoError("'=' 토큰은 현재 위치에서 사용할 수 없습니다.")]
        VB0066,
        #endregion

        #region [  If / Select Case 오류  ]

        [KoError("If문을 닫은 다음에 다른 문법 처리를 시작해야 합니다.")]
        VB0069 = 69,

        /// <summary>ElseIf If문 이후에 사용 가능 오류</summary>
        [KoError("ElseIf문은 If문 이후에 사용 할 수 있습니다.")]
        VB0070,

        /// <summary>Else If문 이후에 사용 가능 오류</summary>
        [KoError("Else문은 If문 이후에 사용 할 수 있습니다.")]
        VB0071,
        /// <summary>Else If문 한줄 사용 오류</summary>
        [KoError("ElseIf문은 한줄 If문으로 사용할 수 없습니다.")]
        VB0072,
        /// <summary>If문 한줄 중복 사용 오류</summary>
        [KoError("If문은 한줄에 중복해서 사용 할 수 없습니다.")]
        VB0073,
        /// <summary>If문 처음에만 사용 가능</summary>
        [KoError("If문은 맨 처음에만 사용 할 수 있습니다.")]
        VB0074,
        /// <summary>If서 Else와 If띄워 사용 불가능</summary>
        [KoError("VBA에서는 Else와 If를 띄워서 사용할 수 없습니다. 대신 ElseIf를 사용해보세요.")]
        VB0075,
        /// <summary>Select에서 Case누락</summary>
        [KoError("Case가 필요합니다.")]
        VB0076,

        [KoError("Then이 누락되었습니다. If문의 마지막에는 Then이 필요합니다.")]
        VB0077,

        [KoError("조건문 뒤에는 바로 연산자가 올 수 없습니다.")]
        VB0078,

        [KoError("End If로 닫을 If문을 찾을 수 없습니다.")]
        VB0079,


        #endregion

        #region [  Do / While 오류  ]
        
        [KoError("Do문은 맨 처음에만 오거나 Exit Do의 형태로만 사용 할 수 있습니다.")]
        VB0080 = 80,
        [KoError("Until문은 Do Until로만 사용 할 수 있습니다.")]
        VB0081,
        [KoError("Wend문은 While문을 닫을 때만 사용 할 수 있습니다.")]
        VB0082,

        /// <summary>Do문 이후 식 사용 불가능</summary>
        [KoError("Do문 이후에는 식을 사용 할 수 없습니다.")]
        VB0083,

        /// <summary>Loop문 이후 식 사용 불가능</summary>
        [KoError("Loop은 처음에만 사용 할 수 있습니다.")]
        VB0084,

        /// <summary>Loop문 이후 식 사용 불가능</summary>
        [KoError("Wend문 이후에는 식을 사용 할 수 없습니다.")]
        VB0085,

        [KoError("Wend로 닫을 While문을 찾을 수 없습니다.")]
        VB0086,

        [KoError("Loop으로 닫을 Do문을 찾을 수 없습니다.")]
        VB0087,

        #endregion



        #region [  파라미터 오류  ]

        [KoError("ParamArray는 Optional과 함께 사용 할 수 없습니다.")]
        VB0090 = 90,
        [KoError("Optional은 ParamArray와 함께 사용 할 수 없습니다.")]
        VB0091,
        [KoError("ParamArray는 한번만 사용할 수 있습니다.")]
        VB0092,
        [KoError("Optional이나 ParamArray 이후 절에는 ByVal이나 ByRef가 사용될 수 없습니다.")]
        VB0093,
        [KoError("ParamArray는 Variant의 배열이여야 합니다.")]
        VB0094,
        [KoError("Optional을 사용할 시 초기값이 반드시 들어가야 합니다.")]
        VB0095,
        /// <summary>선언문 (Dim, Sub, Function) 오류</summary>
        [KoError("파라미터 인식에는 Dim을 사용 할 수 없습니다. 대신 ByVal, ByRef, ParamArray를 사용해주세요.")]
        VB0096,
        /// <summary>파라미터 식별자 필요</summary>
        [KoError("파라미터 식별자가 필요합니다.")]
        VB0097,
        /// <summary>파라미터 쉼표 필요</summary>
        [KoError("쉼표가 필요하지 않습니다.")]
        VB0098,
        /// <summary>파라미터 쉼표 필요하지 않음</summary>
        [KoError("쉼표가 필요합니다.")]
        VB0099,
        /// <summary>파라미터 Optional 타입 오류</summary>
        [KoError("Optional은 기본값을 가진 내부 형식이나 Variant이어야 합니다.")]
        VB0100,
        /// <summary>파라미터 괄호 오류</summary>
        [KoError("파라미터 배열 선언은 괄호 한번만 열어야 합니다.")]
        VB0101,
        [KoError("ByVal, ByRef, ParamArray, Optional 등은 파라미터 내부에서만 사용 할 수 있습니다.")]
        VB0102,
        [KoError("Optional을 제외한 나머지에는 초기값이 들어갈 수 없습니다.")]
        VB0103,

        [KoError("현 위치에서는 String 값을 사용할 수 없습니다.")]
        VB0104,

        [KoError("열린 String은 새줄이 시작하기 전에 닫혀야 합니다.")]
        VB0105,

        [KoError("열린 String은 새줄이 시작하기 전에 닫혀야 합니다.")]
        VB0106,

        #endregion

        #region [  선언문 오류  ]

        /// <summary>선언문 (Dim, Sub, Function) 오류</summary>
        [KoError("선언문은 엑세서를 생략하고 쓰거나 엑세서 이후에만 사용 할 수 있습니다.")]
        VB0120 = 120,

        /// <summary>식별자가 필요합니다.</summary>
        [KoError("식별자가 필요합니다.")]
        VB0121,
        /// <summary>식이 필요합니다.</summary>
        [KoError("식이 필요합니다.")]
        VB0122,
        /// <summary>식별자/선언자 필요</summary>
        [KoError("식별자 또는 Sub/Function/Property가 와야 합니다.")]
        VB0123,
        /// <summary>Type 필요</summary>
        [KoError("타입이 필요합니다.")]
        VB0124,
        /// <summary>ByVal, ByRef, ParamArray, Optional, ')' 이 필요합니다.</summary>
        [KoError("ByVal, ByRef, ParamArray, Optional 또는 ')' 이 필요합니다.")]
        VB0125,

        /// <summary>Exit .. Sub/Function/Property/Do/For</summary>
        [CanReplace(1)]
        [KoError("Exit %1로 닫을 %1문을 찾을 수 없습니다.")]
        VB0127,

        #endregion

        #region [  라인 위치 오류  ]

        /// <summary>Set/Get/Let 위치 오류</summary>
        [CanReplace(1)]
        [KoError("%1 컨텍스트는 현재 위치에 올 수 없습니다.")]
        VB0130 = 130,

        /// <summary>Property + Get/Let/Set</summary>
        [KoError("Property 절 뒤에는 Get이나 Let 또는 Set만 올 수 있습니다.")]
        VB0131,

        /// <summary>Function Nest 사용 불가능</summary>
        [KoError("Function은 메소드에 네스트해 사용할 수 없습니다.")]
        VB0132,

        /// <summary>Sub Nest 사용 불가능</summary>
        [KoError("Sub는 메소드에 네스트해 사용할 수 없습니다.")]
        VB0133,

        /// <summary>Property Nest 사용 불가능</summary>
        [KoError("Property는 메소드에 네스트해 사용할 수 없습니다.")]
        VB0134,

        /// <summary>Function Enum 안에 사용 불가능</summary>
        [KoError("Function은 Enum 안에 쓸 수 없습니다.")]
        VB0135,

        /// <summary>Sub Enum 안에 사용 불가능</summary>
        [KoError("Sub는 Enum 안에 쓸 수 없습니다.")]
        VB0136,

        /// <summary>Property Enum 안에 사용 불가능</summary>
        [KoError("Property는 Enum 안에 쓸 수 없습니다.")]
        VB0137,

        /// <summary>End Function 단독 사용 불가능</summary>
        [KoError("End Function은 Function 뒤에만 올 수 있습니다.")]
        VB0138,

        /// <summary>End Sub 단독 사용 불가능</summary>
        [KoError("End Sub는 Sub 뒤에만 올 수 있습니다.")]
        VB0139,

        /// <summary>End Property 단독 사용 불가능</summary>
        [KoError("End Property는 Property 뒤에만 올 수 있습니다.")]
        VB0140,

        
        /// <summary>While/If/Do/For 프로시져 내부에서만 사용 가능 명시</summary>
        [CanReplace(1)]
        [KoError("%1은(는) 프로시져 내부에서만 사용할 수 있습니다.")]
        VB0141,

        /// <summary>프로시져 내부에서 사용 불가능 명시</summary>
        [CanReplace(1)]
        [KoError("%1은(는) 프로시져 내부에서 사용 할 수 없습니다.")]
        VB0142,

        [KoError("프로시져 내부의 변수 선언은 Dim 키워드로 선언해야 합니다.")]
        VB0143,

        [KoError("프로시져 내부에서는 상수를 정의할 수 없습니다.")]
        VB0144,

        [KoError("전역 변수는 프로시져가 시작하기 전 위치까지만 사용 할 수 있습니다.")]
        VB0145,

        #endregion

        #region [  배열/As 오류  ]

        /// <summary>배열 크기 뒤에는 상수 필요</summary>
        [KoError("배열 크기 선언에는 상수가 와야합니다")]
        VB0150 = 150,

        /// <summary>As 절은 식별자 다음에만 올 수 있습니다.</summary>
        [KoError("As 절은 식별자 다음에만 올 수 있습니다.")]
        VB0151,

        /// <summary>As가 와야 합니다.</summary>
        [KoError("As가 와야 합니다.")]
        VB0152,

        [KoError("ReDim은 맨 처음에만 올 수 있습니다.")]
        VB0153,

        [KoError("ReDim 뒤에는 식별자가 필요합니다.")]
        VB0154,

        [KoError("ReDim의 배열 크기 부분은 비워둘수 없습니다.")]
        VB0155,


        #endregion



        #region [  문자열 오류  ]

        [KoError("문자열은 닫아야 합니다.")]
        VB0160 = 160,

        #endregion

        #region [  For / For Each 오류  ]

        [KoError("For문은 맨 처음에만 올 수 있습니다.")]
        VB0170 = 170,

        [KoError("Each문은 For문 뒤에 올 수 있습니다.")]
        VB0171,

        [KoError("For문은 End로 닫을 수 없습니다. End For문 대신 Next를 사용하세요.")]
        VB0172,

        [KoError("Next로 닫을 For문을 찾을 수 없습니다. 또는 On Error Resume Next로 사용 가능합니다.")]
        VB0173,

        [KoError("In 키워드는 For Each 문에서만 사용 할 수 있습니다.")]
        VB0174,

        
        #endregion

        #region [  On Error [Goto Label / Resume Next]  ]

        [KoError("On 키워드는 맨처음에만 올 수 있습니다.")]
        VB0180 = 180,

        [KoError("Error 키워드는 On 이후에만 올 수 있습니다.")]
        VB0181,

        [KoError("Goto 키워드는 On Error 이후에만 올 수 있습니다.")]
        VB0182,

        [KoError("Resume 키워드는 On Error 이후에만 올 수 있습니다.")]
        VB0183,

        [KoError("On Error Resume Next 이후에는 아무것도 올 수 없습니다.")]
        VB0184,

        [KoError("On Error GoTo Label 이후에는 아무것도 올 수 없습니다.")]
        VB0185,

        [KoError("On 키워드 이후에는 Error이 나와야 합니다.")]
        VB0186,

        [KoError("On Error이후에는 GoTo Label 또는 Resume Next가 나와야 합니다.")]
        VB0187,

        [KoError("On Error GoTo 이후에는 라벨 이름이 나와야 합니다.")]
        VB0188,

        [KoError("On Error Resume 이후에는 Next이 나와야 합니다.")]
        VB0189,

        [KoError("On Error Goto 이후에는 예약어가 올 수 없습니다.")]
        VB0190,

        #endregion


        #region [  Declare(API) 선언 오류  ]

        /// <summary>Declare는 선언자 앞에 올 수 없습니다.</summary>
        [KoError("Declare은 선언자 앞에 올수 없습니다.")]
        VB0200 = 200,

        [KoError("Declare문은 Sub나 Function으로만 정의가 가능합니다.")]
        VB0201,

        [KoError("Declare Sub/Function Name 뒤에는 Lib가 와야 합니다.")]
        VB0202,

        [KoError("Lib 뒤에는 String으로 변환 가능한 식이 필요합니다.")]
        VB0203,

        [KoError("Alias가 필요합니다.")]
        VB0204,

        [KoError("Alias 뒤에는 String으로 변환 가능한 식이 필요합니다.")]
        VB0205,

        [KoError("Lib는 Declare Sub/Function Name 뒤에 나와야 합니다.")]
        VB0206,

        [KoError("Alias는 Declare Sub/Function Name Lib LibName 뒤에 나와야 합니다.")]
        VB0207,

        [CanReplace(1)]
        [KoError("%1(은)는 중복해서 사용 할 수 없습니다.")]
        VB0208,

        [KoError("Sub나 Function이 필요합니다.")]
        VB0209,
        #endregion

        #region [  Nest 오류  ]

        [KoError("Function 블록이 닫히지 않았습니다. End Function로 닫아주세요.")]
        VB0210 = 210,

        [KoError("Sub 블록이 닫히지 않았습니다. End Sub로 닫아주세요.")]
        VB0211,

        [KoError("Property 블록이 닫히지 않았습니다. End Property로 닫아주세요.")]
        VB0212,

        [KoError("Select Case 블록이 닫히지 않았습니다. End Select로 닫아주세요.")]
        VB0213,

        [KoError("If 블록이 닫히지 않았습니다. End If로 닫아주세요.")]
        VB0214,

        [KoError("Do 블록이 닫히지 않았습니다. Loop으로 닫아주세요.")]
        VB0215,

        [KoError("Do While 블록이 닫히지 않았습니다. Loop으로 닫아주세요.")]
        VB0216,

        [KoError("Do Until 블록이 닫히지 않았습니다. Loop으로 닫아주세요.")]
        VB0217,

        [KoError("Enum 블록이 닫히지 않았습니다. End Enum으로 닫아주세요.")]
        VB0218,

        [KoError("For 블록이 닫히지 않았습니다. Next로 닫아주세요.")]
        VB0219,

        [KoError("For Each 블록이 닫히지 않았습니다. Next로 닫아주세요.")]
        VB0220,

        #endregion


        #region [  With 오류  ]

        [KoError("End With는 With문 다음에만 올 수 있습니다.")]
        VB0230 = 230,

        #endregion


        #region [  Option 오류  ]


        [KoError("Option은 첫번째 단어에만 올 수 있습니다.")]
        VB0240 = 240,

        [KoError("Option 이후에는 Base, Compare, Explicit 또는 Private가 필요합니다.")]
        VB0241,
        [KoError("Option 이후에는 Base, Compare, Explicit 또는 Private만 와야합니다.")]
        VB0242,
        [KoError("Option은 첫번째 단어에만 올 수 있습니다.")]
        VB0243,

        // Option Compare {Text|Binary}
        [KoError("Option Compare 다음에는 Text나 Binary가 와야 합니다.")]
        VB0245,
        [KoError("Option Compare 다음에는 Text나 Binary만 와야 합니다.")]
        VB0246,

        [KoError("Option Base 다음에는 0이나 1이 와야 합니다.")]
        VB0247,

        [KoError("Option Base 다음에는 0이나 1만 와야 합니다.")]
        VB0248,

        // Option Explicit
        [KoError("Explicit은 Option 뒤에만 올 수 있습니다.")]
        VB0253,

        [KoError("Compare은 Option 뒤에만 올 수 있습니다.")]
        VB0254,


        [KoError("Option Compare 다음에는 Text나 Binary만 와야 합니다.")]
        VB0255,

        [KoError("Text는 Option Compare 다음에만 올 수 있습니다.")]
        VB0256,

        [KoError("Binary는 Option Compare 다음에만 올 수 있습니다.")]
        VB0257,

        // Option Private Module

        [KoError("Option Private 다음에는 Module이 와야 합니다.")]
        VB0258,

        [KoError("Option Private 다음에는 Module만 와야 합니다.")]
        VB0259,

        [KoError("Option Private Module 이후에는 다른 식이 올 수 없습니다.")]
        VB0260,

        [KoError("Base는 Option 뒤에만 올 수 있습니다.")]
        VB0261,

        [KoError("Option Explicit 이후에는 문을 사용 할 수 없습니다.")]
        VB0262,

        #endregion
    }
}
