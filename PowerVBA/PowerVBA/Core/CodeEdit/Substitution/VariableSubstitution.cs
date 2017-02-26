using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Core.CodeEdit.Substitution.Base;
using PowerVBA.Core.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static PowerVBA.Core.CodeEdit.Convert.StringConverter;
using PowerVBA.Global.Regex;


namespace PowerVBA.Core.CodeEdit.Substitution
{
    class VariableSubstitution : BaseSubstitution
    {
        public VariableSubstitution (TextEditor editor) : base(editor) { }
        
        
        public override void Convert()
        {
            DocumentLine line = Editor.Document.GetLineByOffset(Editor.CaretOffset);

            string clonestr = Editor.Text.Clone().ToString();
            string codeText = (clonestr.Substring(line.Offset, Editor.CaretOffset - line.Offset));
            string anotherText = clonestr.Substring(line.Offset + codeText.Length, line.EndOffset - (codeText.Length + line.Offset));

            if (!Regex.IsMatch(anotherText, Pattern.blankCheckPattern)) return;

            int Offset = line.Offset;

            if (Regex.IsMatch(codeText, Pattern.VariableDeclarePattern))
            {
                string TempStr = clonestr;
                Match m = Regex.Match(codeText, Pattern.VariableDeclarePattern);

                string Accessor = m.Groups[1].Value;
                string Name = m.Groups[3].Value;
                string Type = m.Groups[6].Value;

                // 변수 선언 부분
                string DeclarePart = $"{GetIndentation()}{ConvertAccessor(Accessor)} {Name} As {Type}\r\n{GetIndentation()}";

                //TempStr = TempStr.Change(line.Offset, line.Length, DeclarePart);

                Editor.TextArea.Document.Replace(line.Offset, line.Length, DeclarePart);
                //Editor.TextArea.Caret.Offset = Offset + DeclarePart.Length;

                Handled = true;
            }
        }
    }
}
