using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using System.Text.RegularExpressions;
using PowerVBA.RegexPattern;
using PowerVBA.Core.AvalonEdit.Substitution.Base;
using static PowerVBA.Core.Convert.StringConverter;
using System;
using System.Collections;
using System.Collections.Generic;
using ICSharpCode.AvalonEdit.Editing;

namespace PowerVBA.Core.AvalonEdit.Substitution
{
    class MethodSubstitution : BaseSubstitution
    {
        
        public MethodSubstitution(TextArea tArea) : base(tArea) { }
        
        public override void Convert()
        {

            int CaretOffset = TextArea.Caret.Offset;
            
            DocumentLine line = TextArea.Document.GetLineByOffset(CaretOffset);
            
            string clonestr = TextArea.Document.Text.ToString();
            string codeText = (clonestr.Substring(line.Offset, CaretOffset - line.Offset));
            string anotherText = clonestr.Substring(line.Offset + codeText.Length, line.EndOffset - (codeText.Length + line.Offset));

            if (!System.Text.RegularExpressions.Regex.IsMatch(anotherText, Pattern.BlankCheckPattern)) return;
            
            int Offset = line.Offset;

            if (System.Text.RegularExpressions.Regex.IsMatch(codeText.Trim(), Pattern.LineStartPattern))
            {
                Match m = System.Text.RegularExpressions.Regex.Match(codeText.Trim(), Pattern.LineStartPattern);

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


                TextArea.Document.Replace(line.Offset, line.Length, 
                                                 sMethodPart + "\r\n" + GetIndentation() + "\t" + "\r\n" + eMethodPart);

                TextArea.Caret.Offset = Offset + sMethodPart.Length + 3;

                Handled = true;
            }
        }
    }
}
