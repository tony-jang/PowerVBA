using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public struct CodeData
    {
        #region [  Basis  ]

        /// <summary>
        /// 첫째 줄인지 여부를 가져옵니다.
        /// </summary>
        public bool IsFistNonWs { get; set; }

        /// <summary>
        /// 주석 여부를 가져옵니다.
        /// </summary>
        public bool IsInComment { get; set; }

        /// <summary>
        /// string 내부 인지를 가져옵니다.
        /// </summary>
        public bool IsInString { get; set; }
        
        /// <summary>
        /// 그대로의 String("와 같은) 인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInVerbatimString { get; set; }
        
        /// <summary>
        /// 전처리기 지시문인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInPreprocessorDirective { get; set; }

        #endregion


        #region [  Parameter  ]

        /// <summary>
        /// 파라미터 선언 내부인지를 가져옵니다.
        /// </summary>
        public bool IsInParameters { get; set; }

        /// <summary>
        /// 파라미터 ParamArray 이후 절인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool AfterParamArray { get; set; }

        /// <summary>
        /// 파라미터 Optional 이후 절인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool AfterOptional { get; set; }

        #endregion


        /// <summary>
        /// 현재 변수 선언 중인지 여부를 나타냅니다. 확실치 않을때도 false를 반환합니다.
        /// </summary>
        public bool IsVarDeclaring { get; set; }

        /// <summary>
        /// 현재 Function 선언 중인지 여부를 나타냅니다. 확실치 않을때도 false를 반환합니다.
        /// </summary>
        public bool IsFuncDeclaring { get; set; }

        /// <summary>
        /// 현재 Sub 선언 중인지 여부를 나타냅니다. 확실치 않을때도 false를 반환합니다.
        /// </summary>
        public bool IsSubDeclaring { get; set; }

        /// <summary>
        /// Public이나 Private 같은 엑세서 이후인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool AfterAccessor { get; set; }
        
        /// <summary>
        /// Sub나 Function 같은 선언자 이후인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool AfterDeclarator { get; set; }

        public bool AfterArray { get; set; }



        #region [  Do Until/While  ]

        public bool AfterDo { get; set; }
        public bool AfterUntil { get; set; }
        public bool AfterWhile { get; set; }
        public bool AfterLoop { get; set; }

        #endregion
        
        /// <summary>
        /// 식별자 이후인지에 대한 여부를 가져옵니다. 이 이후에는 보통 아무것도 없거나, As 키워드가 나와야 합니다.
        /// </summary>
        public bool AfterIdentifier { get; set; }

        #region [  For  ]

        /// <summary>
        /// For 이후인지에 대한 여부를 가져옵니다. 이 이후에는 식별자가 나와야 합니다.
        /// </summary>
        public bool AfterFor { get; set; }
        /// <summary>
        /// For Each 이후인지에 대한 여부를 가져옵니다. 이 이후에는 식별자가 나와야 합니다.
        /// </summary>
        public bool AfterForEach { get; set; }

        #endregion

        /// <summary>
        /// 괄호 안인지에 대한 여부입니다.
        /// </summary>
        public bool IsInBracket { get; set; }

        #region [  If  ]
        /// <summary>
        /// If 절 이후인지에 대한 여부를 가져옵니다. 이후에는 Expression이 나와야 합니다.
        /// </summary>
        public bool AfterIf { get;  set; }
        /// <summary>
        /// ElseIf 절 이후인지에 대한 여부를 가져옵니다. 이후에는 Expression이 나와야 합니다.
        /// </summary>
        public bool AfterElseIf { get; set; }

        #endregion
        
        #region [  Select Case  ]

        /// <summary>
        /// Select 절 이후인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool AfterSelect { get; set; }
        public bool AfterCase { get; set; }

        #endregion




        #region [  Keyword  ]

        /// <summary>
        /// Else 절 이후인지에 대한 여부를 가져옵니다. 이후에는 아무것도 없거나, 
        /// If 한줄 절일때는 Statement를 사용해야 합니다.
        /// </summary>
        public bool AfterElse { get; set; }

        public bool AfterExit { get; set; }

        /// <summary>
        /// As 이후인지에 대한 여부를 가져옵니다. 이 이후에는 타입이 나와야 합니다.
        /// </summary>
        public bool AfterAs { get; set; }


        /// <summary>
        /// End 절 이후인지를 가져옵니다 이 이후에는 아무것도 없거나 
        /// Sub, Function, Type, If, Select등이 나와야 합니다.
        /// </summary>
        public bool AfterEnd { get; set; }


        #endregion


        /// <summary>
        /// Property 절 이후인지에 대한 여부를 가져옵니다. 이후에는 Expression이 나와야 합니다.
        /// </summary>
        public bool AfterProperty { get; set; }
    }
}