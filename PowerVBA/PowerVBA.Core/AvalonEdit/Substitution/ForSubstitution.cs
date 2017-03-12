using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;
using PowerVBA.Core.AvalonEdit.Substitution.Base;
using PowerVBA.RegexPattern;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static PowerVBA.Core.Convert.StringConverter;
using System.Threading.Tasks;
using PowerVBA.Global.RegexExpressions;

namespace PowerVBA.Core.AvalonEdit.Substitution
{
    class ForSubstitution : BaseSubstitution
    {

        public ForSubstitution(TextArea textarea) : base(textarea) { }
        public enum ForTypes
        {
            None,
            Standard,
            Step,
            Standard_Ex,
            Step_Ex
        }
        public override void Convert()
        {
            int CaretOffset = TextArea.Caret.Offset;

            DocumentLine line = TextArea.Document.GetLineByOffset(CaretOffset);
            
            string codeText = (TextArea.Document.Text.Substring(line.Offset, CaretOffset - line.Offset));

            int Offset = line.Offset;

            ForTypes forTypes = ForTypes.None;

            string VarName, Type, StartPos, EndPos, Step;

            string sForPart = "", eForPart = "";

            Match m;

            if (Regex.IsMatch(codeText, CodePattern.ForStandard))
            {
                m = Regex.Match(codeText, CodePattern.ForStandard);

                VarName = m.Groups[1].Value;
                StartPos = m.Groups[2].Value;
                EndPos = m.Groups[3].Value;

                forTypes = ForTypes.Standard;
                
                sForPart = $"For {VarName} = {StartPos} To {EndPos}";
                eForPart = "Next";
            }
            else if (Regex.IsMatch(codeText, CodePattern.ForStep))
            {
                m = Regex.Match(codeText, CodePattern.ForStep);

                VarName = m.Groups[1].Value;
                StartPos = m.Groups[2].Value;
                EndPos = m.Groups[3].Value;
                Step = m.Groups[4].Value;

                forTypes = ForTypes.Step;
                
                sForPart = $"For {VarName} = {StartPos} To {EndPos} Step {Step}";
                eForPart = "Next";
            }
            else if(Regex.IsMatch(codeText, CodePattern.ForStandard_Ex))
            {
                m = Regex.Match(codeText, CodePattern.ForStandard_Ex);

                VarName = m.Groups[1].Value;
                Type = m.Groups[2].Value;
                StartPos = m.Groups[3].Value;
                EndPos = m.Groups[4].Value;

                forTypes = ForTypes.Standard_Ex;

                sForPart = $"For {VarName} As {Type} = {StartPos} To {EndPos}";
                eForPart = "Next";
            }
            else if (Regex.IsMatch(codeText, CodePattern.ForStep_Ex))
            {
                m = Regex.Match(codeText, CodePattern.ForStep_Ex);

                VarName = m.Groups[1].Value;
                Type = m.Groups[2].Value;
                StartPos = m.Groups[3].Value;
                EndPos = m.Groups[4].Value;
                Step = m.Groups[5].Value;

                forTypes = ForTypes.Step_Ex;

                sForPart = $"For {VarName} As {Type} = {StartPos} To {EndPos} Step {Step}";
                eForPart = "Next";
            }

            if (forTypes != ForTypes.None)
            {
                TextArea.Document.Replace(line.Offset, line.Length, sForPart + "\r\n\t" + "\r\n" + eForPart);

                TextArea.Caret.Offset = Offset + sForPart.Length + 3;

                Handled = true;
            }
        }

    }
}
