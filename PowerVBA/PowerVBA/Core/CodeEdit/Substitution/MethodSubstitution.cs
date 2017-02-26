using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using System.Text.RegularExpressions;
using PowerVBA.Core.Extension;
using PowerVBA.Global.Regex;
using static PowerVBA.Core.CodeEdit.Convert.StringConverter;
using PowerVBA.Core.CodeEdit.Substitution.Base;
using System;
using ICSharpCode.AvalonEdit.CodeCompletion;
using System.Collections;
using System.Collections.Generic;

namespace PowerVBA.Core.CodeEdit.Substitution
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


            List<ICompletionData> data;
            

            int Offset = line.Offset;

            if (Regex.IsMatch(codeText.Trim(), Pattern.lineStartPattern))
            {
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


                Editor.TextArea.Document.Replace(line.Offset, line.Length, 
                                                 sMethodPart + "\r\n" + GetIndentation() + "\t" + "\r\n" + eMethodPart);

                Editor.TextArea.Caret.Offset = Offset + sMethodPart.Length + 3;

                Handled = true;
            }
        }
    }
}
