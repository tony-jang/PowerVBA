using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Core.Project;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Parser
{
    class LineParser
    {

        public LineParser(string text, CodeParser mainParser, AnchorSegment segment)
        {
            LineText = text;
            Parent = mainParser;
            Segment = segment;
        }

        string LineText;
        CodeParser Parent;
        AnchorSegment Segment;

        public ILineInfo Parse()
        {
            string parseText = LineText.TrimStart();
            ILineInfo value = null;
            
            if (parseText.StartsWith("'"))
            {
                // 주석 인식
                value = new SingleLineInfo(Segment.Length, Segment.Offset, LineType.Remark);
            }
            else if (parseText.StartsWith("Dim"))
            {

                
                // 변수 선언 인식
                value = new SingleLineInfo(Segment.Length, Segment.Offset, LineType.GlobalVariable);
            }
            return value;
        }
    }
}
