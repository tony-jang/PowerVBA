using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using System.Text.RegularExpressions;
using PowerVBA.Core.Extension;
using PowerVBA.Global.Regex;
using static PowerVBA.Core.AvalonEdit.Convert.StringConverter;
using PowerVBA.Core.AvalonEdit.Substitution.Base;
using System;

namespace PowerVBA.Core.AvalonEdit.Substitution
{
    class MethodSubstitution : BaseSubstitution
    {
        
        public MethodSubstitution(TextEditor editor) : base(editor) { }


        #region [  Pattern  ]
        
        

        #endregion


        public override void Convert()
        {

            int d = Editor.CaretOffset;
            
            DocumentLine line = Editor.Document.GetLineByOffset(Editor.CaretOffset);

            
            string clonestr = Editor.Text.Clone().ToString();
            string codeText = (clonestr.Substring(line.Offset, Editor.CaretOffset - line.Offset));
            string anotherText = clonestr.Substring(line.Offset + codeText.Length, line.EndOffset - (codeText.Length + line.Offset));

            if (!Regex.IsMatch(anotherText, Pattern.blankCheckPattern)) return;

            int Offset = line.Offset;

            if (Regex.IsMatch(codeText.Trim(), Pattern.lineStartPattern))
            {
                string TempStr;
                Match m = Regex.Match(codeText.Trim(), Pattern.lineStartPattern);

                string Accessor = m.Groups[1].Value;
                string Type = m.Groups[3].Value;
                string Name = m.Groups[5].Value;
                string parameter = "";

                if (!string.IsNullOrEmpty(m.Groups[7].Value)) parameter = m.Groups[8].Value;

                // 메소드 선언 부분
                string sMethodPart = $"{ConvertAccessor(Accessor)} {ConvertType(Type)} {Name}";
                if (Type != "Type") sMethodPart += $"({parameter})";
                // 메소드 종료 부분
                string eMethodPart = "End " + ConvertType(Type);


                TempStr = Editor.Text.Insert(line.Offset + line.Length, "\r\n" + GetIndentation() + "\t" + "\r\n" + eMethodPart);
                TempStr = TempStr.Change(line.Offset, line.Length, sMethodPart);

                Editor.Text = TempStr;
                Editor.TextArea.Caret.Offset = Offset + sMethodPart.Length + 3;

                Handled = true;
            }
        }
    }
}
