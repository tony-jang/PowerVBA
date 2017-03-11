using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Core.Project;
using PowerVBA.RegexPattern;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Parser
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
            ILineInfo value = new SingleLineInfo(Segment.Length, Segment.Offset, LineType.Unknown);
            
            if (parseText.StartsWith("'"))
            {
                // 주석 인식
                value = new SingleLineInfo(Segment.Length, Segment.Offset, LineType.Remark);
            }
            else if (Regex.IsMatch(parseText, Pattern.pattern4))
            {
                // 변수 선언 인식
                Match m = Regex.Match(parseText, Pattern.pattern4);
                LineType lineType = LineType.Unknown;
                string type1 = m.Groups[2].Value.ToLower();
                //if ()

                value = new SingleLineInfo(Segment.Length, Segment.Offset, LineType.GlobalVariable);
            }
            else if (Regex.IsMatch(parseText, Pattern.pattern3))
            {
                //4
                Match m = Regex.Match(parseText, Pattern.pattern3);

                string type = m.Groups[4].Value.ToLower();
                LineType lineType = LineType.Unknown;

                if (type == "enum") { lineType = LineType.EnumStart; }
                else if (type == "type") { lineType = LineType.TypeStart; }

                value = new SingleLineInfo(Segment.Length, Segment.Offset, lineType);
            }
            
            return value;
        }
    }
}
