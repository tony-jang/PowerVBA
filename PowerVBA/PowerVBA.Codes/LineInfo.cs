using System;
using System.Collections.Generic;
using PowerVBA.Codes.Parsing;

namespace PowerVBA.Codes
{
    /// <summary>
    /// 라인을 넘어가며 네스트 되어 있는 걸 확인합니다.
    /// </summary>
    public class LineInfo
    {

        public LineInfo()
        {
            IsGlobalVarDeclaring = true;
            Foldings = new List<RangeInt>();
        }

        public string FileName { get; set; }

        public List<RangeInt> Foldings { get; set; }

        public Stack<(int, FoldingTypes)> TempLine { get; set; } = new Stack<(int, FoldingTypes)>();

        /// <summary>
        /// 마지막으로 글로벌 변수가 선언된 위치를 확인합니다. 0을 반환하면 처음 부분에 새 줄을 만든 뒤 선언합니다.
        /// </summary>
        public int LastGlobalVarInt { get; set; }

        /// <summary>
        /// Sub나 Function 또는 Property 내부에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInProcedure { get { return IsInSub | IsInFunction | IsInProperty; } }

        public Locator CodeLocator { get; set; } = new Locator();


        #region [  Method  ]

        /// <summary>
        /// Sub 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInSub => CodeLocator.Contains(CodeType.SubBlock);
        /// <summary>
        /// Function 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInFunction => CodeLocator.Contains(CodeType.FunctionBlock);
        /// <summary>
        /// Property 내부에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInProperty => CodeLocator.Contains(CodeType.PropertyBlock);

        #endregion

        #region [  With  ]

        /// <summary>
        /// With문 안에 있는지에 대한 여부입니다. With문 내부에 있으면 오브젝트(개체)의 시작이 .으로 시작해도 자동 인식됩니다.
        /// </summary>
        public bool IsInWith => CodeLocator.Contains(CodeType.WithBlock);

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
        public bool IsInIf => CodeLocator.Contains(CodeType.IfBlock);
        /// <summary>
        /// ElseIf문 안에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInElseIf => CodeLocator.Contains(CodeType.ElseIfBlock);
        /// <summary>
        /// Else문 안에 있는지에 대한 여부입니다.
        /// </summary>
        public bool IsInElse => CodeLocator.Contains(CodeType.ElseBlock);

        #endregion

        #region [  Select  ]

        /// <summary>
        /// Case문이나 Case Else문 안에 있는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsNestInSelect { get { return IsInCaseElse | IsInCase; } }

        /// <summary>
        /// Select Case문 안에 있는지에 대한 여부입니다. 해당 옵션이 true인데 IsInCase가 false라면 Case나 Case Else가 작성되어야 합니다.
        /// </summary>
        public bool IsInSelectCase => CodeLocator.Contains(CodeType.SelectCaseBlock);

        /// <summary>
        /// Case문안에 있는지에 대한 여부입니다. 보통은 이 옵션이 True라면 IsInSelectCase도 True여야 합니다.
        /// </summary>
        public bool IsInCase => CodeLocator.Contains(CodeType.CaseBlock);

        /// <summary>
        /// Case Else문 안에 있는지에 대한 여부입니다. 보통은 이 옵션이 True라면 IsInSelectCase도 True여야 합니다.
        /// </summary>
        public bool IsInCaseElse => CodeLocator.Contains(CodeType.CaseElseBlock);

        #endregion


        #region [  For/For Each  ]
        
        public bool IsInFor => CodeLocator.Contains(CodeType.ForBlock);
        public bool IsInForEach => CodeLocator.Contains(CodeType.ForEachBlock);

        #endregion
        
        public bool IsInWhile => CodeLocator.Contains(CodeType.WhileBlock);
        public bool IsInDo => CodeLocator.Contains(CodeType.DoBlock);
        public bool IsInDoWhile => CodeLocator.Contains(CodeType.DoWhileBlock);
        public bool IsInDoUntil => CodeLocator.Contains(CodeType.DoUntilBlock);
        public bool IsInEnum => CodeLocator.Contains(CodeType.EnumBlock);
        public bool IsGlobalVarDeclaring { get; set; }
        public bool IsInType => CodeLocator.Contains(CodeType.TypeBlock);
        
        public Queue<(CodeType, int)> codes = new Queue<(CodeType, int)>();

        /// <summary>
        /// 라인에서 처리된 작업을 예약합니다.
        /// </summary>
        public void Reserving(CodeType type, int Line)
        {
            codes.Enqueue((type, Line));
        }

        public void CancelReserve()
        {
            if (codes.Count > 0)
                codes.Dequeue();
        }

        /// <summary>
        /// 예약해두었던 작업을 모두 실행합니다.
        /// </summary>
        public void AllProcessing()
        {
            while (codes.Count != 0)
            {
                var itm = codes.Dequeue();
                CodeLocator.Insert(itm.Item1, itm.Item2);
            };
        }

        public LineInfo Clone()
        {
            return new LineInfo
            {
                IsGlobalVarDeclaring = IsGlobalVarDeclaring,

                FileName = FileName,
                LastGlobalVarInt = LastGlobalVarInt,

                Foldings = Foldings,
                TempLine = TempLine,

                CodeLocator = CodeLocator,
            };
        }

    }
}
