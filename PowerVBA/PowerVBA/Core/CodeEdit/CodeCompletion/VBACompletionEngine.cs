using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.RegexPattern;
using PowerVBA.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace PowerVBA.Core.CodeEdit.CodeCompletion
{
    class VBACompletionEngine
    {
        public class CompletionEngineCache
        {
            public CompletionEngineCache()
            {
                importCompletion = new List<ICompletionData>();
            }
            public List<ICompletionData> importCompletion;
        }

        protected CompletionEngineCache cache;
        protected IDocument document;
        protected int offset;

        /// <summary>
        /// End Of Line Marker
        /// </summary>
        protected string EolMarker = Environment.NewLine;
        protected string IndentStr = "\t";

        public VBACompletionEngine(IDocument document)
        {
            this.document = document;
            cache = new CompletionEngineCache();
            offset = 0;
        }

        public IEnumerable<ICompletionData> GetCompletionData(int offset, bool controlSpace)
        {
            var data = new List<CompletionData>();


            DocumentLine dl = (DocumentLine)document.GetLineByOffset(offset);
            string s = document.Text.Substring(dl.Offset, dl.Length);

            VisibleItem flags = VisibleItem.None;
            CompletionFlag CompFlag = CompletionFlag.None;

            if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern3))
            {
                flags = VisibleItem.None;
                CompFlag = CompletionFlag.MethodNaming;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern1))
            {
                flags = VisibleItem.Declarer;
                CompFlag = CompletionFlag.OnlyNaming;   
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern1_1))
            {
                flags = VisibleItem.Declarer;
                CompFlag = CompletionFlag.Naming;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern1_2))
            {
                flags = VisibleItem.Declarer;
                CompFlag = CompletionFlag.General;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern1_3))
            {
                flags = VisibleItem.Declarer;
                CompFlag = CompletionFlag.As;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern2))
            {
                flags = VisibleItem.Declarer;
                CompFlag = CompletionFlag.General;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(s, Pattern.pattern4))
            {
                flags = VisibleItem.Class | VisibleItem.Type;
                CompFlag = CompletionFlag.DeclareType;
            }

            if (flags.HasFlag(VisibleItem.Class))
            {
                foreach(string compStr in CodeEditor.Classes)
                {
                    var itm = new CompletionData(compStr, compStr, "Class " + compStr, ResourceImage.GetIconImage("ClassIcon"), CompletionFlag.None);
                    if (cache.importCompletion.Where((elem) => elem.Text == itm.Text).Count() < 1)
                    {
                        data.Add(itm);
                        cache.importCompletion.Add(itm);
                    }
                }
            }
            if (flags.HasFlag(VisibleItem.Enum))
            {

            }
            if (flags.HasFlag(VisibleItem.Method))
            {

            }
            if (flags.HasFlag(VisibleItem.Property))
            {

            }
            if (flags.HasFlag(VisibleItem.Type))
            {
                var itms = (new CompletionData[] {
                           new CompletionData("String", "String", "문자들의 모음입니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.None),
                           new CompletionData("Boolean", "Boolean", "참 또는 거짓의 값에 대해서 나타냅니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.None),
                           new CompletionData("Integer", "Integer", "32비트의 부호 있는 숫자를 나타냅니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.None),
                           new CompletionData("Long", "Long", "64비트의 부호 있는 숫자를 나타냅니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.None)});
                // TODO : https://msdn.microsoft.com/ko-kr/library/47zceaw7.aspx 에서 타입 찾아서 추가하기

                foreach (var itm in itms)
                {
                    if (cache.importCompletion.Where((elem) => elem.Text == itm.Text).Count() < 1)
                    {
                        data.Add(itm);
                        cache.importCompletion.Add(itm);
                    }
                }
            }
            if (flags.HasFlag(VisibleItem.Declarer))
            {
                var itms = (new CompletionData[] {
                           new CompletionData("Dim", "Dim", "하나 이상의 변수에 사용할 저장 공간을 선언하고 할당합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.General),

                           new CompletionData("Public", "Public", "선언된 프로그래밍 요소에 대한 액세스를 제한하지 않도록 설정합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.General),

                           new CompletionData("Private", "Private", "프로그래밍 요소를 선언한 모듈, 클래스에서만 해당 프로그래밍 요소를 액세스할 수 있도록 지정합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.General),

                           new CompletionData("As", "As", "선언문의 데이터 형식을 지정합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.As),

                           new CompletionData("Sub", "Sub", "호출 코드에 값을 반환하지 않는 프로시저인 Sub 프로시저를 정의하는 이름, 매개 변수 및 코드를 선언합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.General | CompletionFlag.Naming),

                           new CompletionData("Function", "Function", "호출 코드에 값을 반환하는 Function 프로시저를 정의하는 이름, 매개 변수 및 코드를 선언합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Declarator | CompletionFlag.General | CompletionFlag.Naming),

                           new CompletionData("Class", "Class", "클래스 이름을 선언하고 클래스를 구성하는 변수, 속성 및 메서드의 정의를 지정합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Naming | CompletionFlag.General),

                           new CompletionData("Enum", "Enum", "열거형을 선언하고 열거형의 멤버 값을 정의합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.Naming | CompletionFlag.General),

                           new CompletionData("With","With","단일 개채 또는 구조체를 참조하는 일련의 문을 실행합니다.",
                               ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.General),

                           new CompletionData("End","End", "블록을 종료합니다.",ResourceImage.GetIconImage("DeclaratorIcon"), CompletionFlag.General),

                           new CompletionData("", "<새 이름>", "새 변수를 선언합니다.", null, CompletionFlag.Naming | CompletionFlag.OnlyNaming | CompletionFlag.MethodNaming) });
                
                foreach(var itm in itms)
                {
                    if (itm.Flag.HasFlag(CompFlag))
                    {
                        if (cache.importCompletion.Where((elem) => elem.Text == itm.Text).Count() < 1)
                        {
                            data.Add(itm);
                            cache.importCompletion.Add(itm);
                        }
                    }
                }
            }


            return data;
        }

        public enum VisibleItem
        {
            None = 0,
            Enum = 1,
            Class = 2,
            Property = 4,
            Method = 8,
            Type = 16,
            Declarer = 32
        }
    }
}
