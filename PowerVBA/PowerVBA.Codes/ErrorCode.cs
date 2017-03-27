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
        /// <summary>선언과 동시에 초기화 불가</summary>
        [NotSupported]
        [KoError("VBA에서는 변수 선언과 동시에 초기화를 지원하지 않습니다.")]
        VB0008,
        VB0009,
        VB0010,

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

        #endregion



        /// <summary>특수 절 (1) 사용 불가</summary>
        [CanReplace(1)]
        [KoError("%1 절은 맨 처음에만 올 수 있습니다.")]
        VB0040 = 40,

        #region [  특수문자 사용 불가  ]

        /// <summary>특수문자 사용 불가</summary>
        [CanReplace(1)]
        [KoError("해당 특수문자 '%1'은(는) 사용할 수 없습니다.")]
        VB0041,


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
        [KoError("%1은(는) End로 닫을 수 없습니다.")]
        VB0060 = 60,

        #endregion

        #region [  If 오류  ]

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
        [KoError("ParamArray는 Variant의 배열이여야 합니다.")]
        VB0092,
        [KoError("Optional을 사용할 시 초기 값이 반드시 들어가야 합니다.")]
        VB0093,
        /// <summary>선언문 (Dim, Sub, Function) 오류</summary>
        [KoError("파라미터 인식에는 Dim을 사용 할 수 없습니다. 대신 ByVal, ByRef, ParamArray를 사용해주세요.")]
        VB0094,
        #endregion

        #region [  선언문 오류  ]
        /// <summary>선언문 (Dim, Sub, Function) 오류</summary>
        [KoError("선언문은 엑세서를 생략하고 쓰거나 엑세서 하나 이후에만 사용 할 수 있습니다.")]
        VB0100 = 100,


        
        #endregion

    }
}
