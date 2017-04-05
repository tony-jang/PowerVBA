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
        /// <summary>미지원 문법 - Set</summary>
        [NotSupported]
        [KoError("'Set'은 VBA에서 지원하지 않습니다. 대신 Public Property Set/Let을 이용해보세요.")]
        VB0004,
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
        [KoError("'%1'라는 이름 또는 형식을 찾을 수 없습니다. 선언되어 있는지 확인하세요.")]
        VB0010,

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
        [KoError("식별자 뒤에는 %1와 같은 예약어가 올 수 없습니다.")]
        VB0045 = 45,
        /// <summary>배열 크기 뒤에는 예약어 사용 불가</summary>
        [KoError("배열 크기 선언에는 예약어가 올 수 없습니다.")]
        VB0046,
        

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

        #endregion

        #region [  If / Select Case 오류  ]

        /// <summary>ElseIf If문 이후에 사용 가능 오류</summary>
        [KoError("ElseIf문은 If문 이후에 사용 할 수 있습니다.")]
        VB0070 = 70,

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

        #endregion

        #region [  Do / While 오류  ]
        
        [KoError("Do문은 맨 처음에만 오거나 Exit Do의 형태로만 사용 할 수 있습니다.")]
        VB0080 = 80,
        [KoError("While문은 Do While로만 사용 할 수 있습니다.")]
        VB0081,
        [KoError("Wend문은 While문을 닫을 때만 사용 할 수 있습니다.")]
        VB0082,

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
        #endregion
        
        #region [  Property 오류  ]

        /// <summary>Set/Get/Let 위치 오류</summary>
        [CanReplace(1)]
        [KoError("%1 컨텍스트는 현재 위치에 올 수 없습니다.")]
        VB0130 = 130,

        /// <summary>Property + Get/Let/Set</summary>
        [KoError("Property 절 뒤에는 Get이나 Let 또는 Set만 올 수 있습니다.")]
        VB0131,
        #endregion

        #region [  As 오류  ]

        /// <summary>As 절은 식별자 다음에만 올 수 있습니다.</summary>
        [KoError("As 절은 식별자 다음에만 올 수 있습니다.")]
        VB0140 = 140,

        /// <summary>As가 와야 합니다.</summary>
        [KoError("As가 와야 합니다.")]
        VB0141,

        #endregion

        #region [  배열 오류  ]

        /// <summary>배열 크기 뒤에는 상수 필요</summary>
        [KoError("배열 크기 선언에는 상수가 와야합니다")]
        VB0150 = 150,

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
        #endregion
        


        #region [  식 오류  ]



        #endregion

    }
}
