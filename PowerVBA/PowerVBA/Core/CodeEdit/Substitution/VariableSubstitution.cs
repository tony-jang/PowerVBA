﻿using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PowerVBA.RegexPattern;
using static PowerVBA.Core.Convert.StringConverter;
using ICSharpCode.AvalonEdit.Editing;
using PowerVBA.Core.CodeEdit.Substitution.Base;


namespace PowerVBA.Core.CodeEdit.Substitution
{
    class VariableSubstitution : BaseSubstitution
    {
        public VariableSubstitution (TextArea tArea) : base(tArea) { }
        
        
        public override void Convert()
        {
            int CaretOffset = TextArea.Caret.Offset;

            DocumentLine line = TextArea.Document.GetLineByOffset(CaretOffset);

            string clonestr = TextArea.Document.Text.Clone().ToString();
            string codeText = (clonestr.Substring(line.Offset, CaretOffset - line.Offset));
            string anotherText = clonestr.Substring(line.Offset + codeText.Length, line.EndOffset - (codeText.Length + line.Offset));

            string StarterIndent = Regex.Match(clonestr, @"^(\t| )+").Value;

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
                // TODO : 앞쪽에 Indent 넣기
                string DeclarePart = $"{ConvertAccessor(Accessor)} {Name} As {Type}\r\n{GetIndentation()}";

                //TempStr = TempStr.Change(line.Offset, line.Length, DeclarePart);

                TextArea.Document.Replace(line.Offset, line.Length, DeclarePart);
                //Editor.TextArea.Caret.Offset = Offset + DeclarePart.Length;

                Handled = true;
            }
        }
    }
}
