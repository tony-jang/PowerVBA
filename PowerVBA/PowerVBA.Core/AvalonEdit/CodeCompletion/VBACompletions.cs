using PowerVBA.Core.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace PowerVBA.Core.AvalonEdit.CodeCompletion
{
    static class VBACompletions
    {
        

        private static string nl = Environment.NewLine;
        private static ImageSource DeclaratorImg = ResourceImage.GetIconImage("DeclaratorIcon");

        public static CompletionData Comp_As = new CompletionData("As", "선언 문의 데이터 형식을 지정합니다.", DeclaratorImg);


        #region [  End  ]

        public static CompletionData Comp_End = new CompletionData("End", "바로 실행을 중지합니다.", DeclaratorImg);
        public static CompletionData Comp_EndSelect = new CompletionData("End Select", "Select Case문의 정의를 종료합니다.", DeclaratorImg);
        public static CompletionData Comp_EndSelect_ = new CompletionData("Select", "Select Case문의 정의를 종료합니다.", DeclaratorImg);
        public static CompletionData Comp_EndIf = new CompletionData("End If", "If문의 정의를 종료합니다.", DeclaratorImg);
        public static CompletionData Comp_EndIf_ = new CompletionData("If", "If문의 정의를 종료합니다.", DeclaratorImg);

        #endregion


        #region [  엑세서  ]

        public static CompletionData Comp_Dim = new CompletionData("Dim", "하나 이상의 변수에 사용할 저장 공간을 선언하고 할당합니다.", DeclaratorImg);
        public static CompletionData Comp_Public = new CompletionData("Public", "선언된 프로그래밍 요소에 대한 액세스를 제한하지 않도록 설정합니다.",DeclaratorImg);
        public static CompletionData Comp_Private = new CompletionData("Private", "프로그래밍 요소를 선언한 모듈, 클래스에서만 해당 프로그래밍 요소를 액세스할 수 있도록 지정합니다.", DeclaratorImg);

        #endregion


        #region [  Do While/Until 문  ]

        public static CompletionData Comp_Do = new CompletionData("Do", "Boolean 조건이 True이거나 조건이 True가 될 때까지 문 블록을 반복합니다.", DeclaratorImg);
        public static CompletionData Comp_DoAfterUntil = new CompletionData("Until", "Boolean 조건이 True가 될 때까지 문 블록을 반복합니다.", DeclaratorImg);
        public static CompletionData Comp_DoAfterWhile = new CompletionData("While", "Boolean 조건이 True인 경우 문 블록을 반복합니다.", DeclaratorImg);

        #endregion

        #region [  If 문  ]

        public static CompletionData Comp_If = new CompletionData("If", "식의 갑에 따라 문의 그룹을 조건부로 실행합니다.", DeclaratorImg);
        public static CompletionData Comp_Else = new CompletionData("Else", "If 문에서 이전 조건이 True가 아닌 경우 실행할 문 그룹을 지정합니다.", DeclaratorImg);
        public static CompletionData Comp_ElseIf = new CompletionData("ElseIf", "If문에서 이전 조건 테스트가 실패할 경우 테스트할 조건을 지정합니다.", DeclaratorImg);

        #endregion

        #region [  Select 문  ]

        public static CompletionData Comp_Select = new CompletionData("Select", "식의 값에 따라 여러 문 그룹 중 하나를 실행합니다.", DeclaratorImg);
        public static CompletionData Comp_SelectCase = new CompletionData("Select Case", "식의 값에 따라 여러 문 그룹 중 하나를 실행합니다.", DeclaratorImg);
        public static CompletionData Comp_Case = new CompletionData("Case", "Select Case 문에서 식의 값을 테스트할 대상 값 또는 값 집합을 지정합니다.", DeclaratorImg);
        public static CompletionData Comp_CaseElse = new CompletionData("Case Else", "Select Case 문에서 이전 조건이 True를 반환하지 않을 경우 실행할 문을 지정합니다.", DeclaratorImg);

        #endregion

        #region [  While 문  ]

        public static CompletionData Comp_While = new CompletionData("While", "지정한 조건이 True인 동안 일련의 문을 실행합니다.", DeclaratorImg);


        #endregion
        
        #region [  For/For Each 문  ]
        public static CompletionData Comp_For = new CompletionData("For", "지정한 횟수만큼 반복되는 루프를 지정합니다.", DeclaratorImg);
        public static CompletionData Comp_To = new CompletionData("To", "루프 카운터나 배열 범위, 범위와 일치하는 값의 시작 값과 끝 값을 구분합니다.", DeclaratorImg);
        public static CompletionData Comp_Step = new CompletionData("Step", "각 루프 반복 간에 증가하는 크기를 지정합니다.", DeclaratorImg);

        public static CompletionData Comp_ForEach = new CompletionData("For Each", "컬렉션의 for each 요소에 대해 반복되는 루프를 지정합니다.", DeclaratorImg);
        public static CompletionData Comp_Each = new CompletionData("Each", "컬렉션의 for each 요소에 대해 반복되는 루프를 지정합니다.", DeclaratorImg);
        public static CompletionData Comp_Next = new CompletionData("Next", "루프 변수 값을 통해 반복하는 루프를 종료합니다.", DeclaratorImg);

        #endregion


    }
}