using ICSharpCode.AvalonEdit.Folding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public class LineInfo
    {

        public LineInfo()
        {
            IsGlobalVarDeclaring = true;
            Foldings = new List<RangeInt>();
        }

        public string FileName { get; set; }

        public List<RangeInt> Foldings { get; set; }


        public Stack<(int, FoldingTypes)> TempLine = new Stack<(int, FoldingTypes)>();

        /// <summary>
        /// 마지막으로 글로벌 변수가 선언된 위치를 확인합니다. 0을 반환하면 처음 부분에 새 줄을 만든 뒤 선언합니다.
        /// </summary>
        public int LastGlobalVarInt { get; set; }

        /// <summary>
        /// Sub나 Function 또는 Property 내부에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInProcedure { get { return IsInSub | IsInFunction | IsInProperty; } }

        #region [  Method  ]

        /// <summary>
        /// Sub 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInSub { get; set; }
        /// <summary>
        /// Function 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInFunction { get; set; }
        /// <summary>
        /// Property 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInProperty { get; set; }

        #endregion

        #region [  With  ]

        /// <summary>
        /// With문 안에 있는지에 대한 여부입니다. With문 내부에 있으면 오브젝트(개체)의 시작이 .으로 시작해도 자동 인식됩니다.
        /// </summary>
        public bool IsInWith { get; set; }

        #endregion


        /// <summary>
        /// If문이나 Select문안에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInConditional { get { return IsNestInIf | IsNestInSelect; } }


        #region [  If  ]
        /// <summary>
        /// If문 내부에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInIf { get { return IsInIf | IsInElseIf | IsInElse; } }

        /// <summary>
        /// If문 안에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInIf { get; set; }
        /// <summary>
        /// ElseIf문 안에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInElseIf { get; set; }
        /// <summary>
        /// Else문 안에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInElse { get; set; }

        #endregion

        #region [  Select  ]

        /// <summary>
        /// Case문이나 Case Else문 안에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInSelect { get { return IsInCaseElse | IsInCase; } }

        /// <summary>
        /// Select Case문 안에 있는지에 대한 여부입니다. 해당 옵션이 true인데 IsInCase가 false라면 Case나 Case Else가 작성되어야 합니다.
        /// </summary>
        public bool IsInSelectCase { get; set; }

        /// <summary>
        /// Case문안에 있는지에 대한 여부입니다. 보통은 이 옵션이 True라면 IsInSelectCase도 True여야 합니다.
        /// </summary>
        public bool IsInCase { get; set; }

        /// <summary>
        /// Case Else문 안에 있는지에 대한 여부입니다. 보통은 이 옵션이 True라면 IsInSelectCase도 True여야 합니다.
        /// </summary>
        public bool IsInCaseElse { get; set; }

        #endregion


        #region [  For/For Each  ]



        public bool IsInFor { get; set; }
        
        public bool IsInForEach { get; set; }

        #endregion


        public bool IsInWhile { get; set; }

        public bool IsInDo { get; set; }

        public bool IsInDoWhile { get; set; }
        public bool IsInDoUntil { get; set; }


        public bool IsInEnum { get; set; }
        public bool IsGlobalVarDeclaring { get; set; }
        public bool IsInType { get; internal set; }

        public LineInfo Clone()
        {
            return new LineInfo
            {
                IsInCase = IsInCase,
                IsInCaseElse = IsInCaseElse,

                IsInDo = IsInDo,
                IsInDoUntil = IsInDoUntil,
                IsInDoWhile = IsInDoWhile,
                IsInWhile = IsInWhile,
                IsInType = IsInType,

                IsInIf = IsInIf,
                IsInElse = IsInElse,
                IsInElseIf = IsInElseIf,

                IsInEnum = IsInEnum,

                IsInFor = IsInFor,
                IsInForEach = IsInForEach,

                IsInFunction = IsInFunction,
                IsInSub = IsInSub,

                IsInProperty = IsInProperty,
                IsInSelectCase = IsInSelectCase,

                IsInWith = IsInWith,

                IsGlobalVarDeclaring = IsGlobalVarDeclaring,

                FileName = FileName,
                LastGlobalVarInt = LastGlobalVarInt,

                Foldings = Foldings,
                TempLine = TempLine,
            };
        }
    }
}
