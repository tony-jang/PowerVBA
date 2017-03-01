using ICSharpCode.AvalonEdit;
using PowerVBA.Core.Error;
using PowerVBA.Global.Regex;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Parser
{
    class CodeParser
    {
        public CodeParser(TextEditor editor, List<CodeError> errors)
        {
            this.Editor = editor;
            this.Errors = errors;

            
        }
        public TextEditor Editor { get; }
        public List<CodeError> Errors { get; }

        public void Seek()
        {
            // 라인 분석 (문법적 선언 오류)
            foreach (string line in Editor.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None))
            {
                string code = line.Trim();

                Match m;
                if (Regex.IsMatch(code, Pattern.g_lineStartPattern))
                {
                    m = Regex.Match(code, Pattern.g_lineStartPattern);

                    

                }
            }

        }
    }
}
