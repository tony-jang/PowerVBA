using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    struct NotHandledLine
    {
        public NotHandledLine(string FileName, string CodeLine, (RangeInt, int) Lines)
        {
            this.FileName = FileName;
            this.CodeLine = CodeLine;
            this.Lines = Lines;
        }
        public string FileName { get; set; }
        public string CodeLine { get; set; }
        public (RangeInt, int) Lines { get; set; }


        public static NotHandledLine Empty = new NotHandledLine("", "", (-1, -1));


        public static bool operator !=(NotHandledLine v1, NotHandledLine v2)
        {
            return (v1.CodeLine != v2.CodeLine &&
                    v1.FileName != v2.FileName &&
                    v1.Lines.Item1.StartInt != v2.Lines.Item1.StartInt &&
                    v1.Lines.Item1.EndInt != v2.Lines.Item1.EndInt &&
                    v1.Lines.Item2 != v2.Lines.Item2);
        }
        public static bool operator ==(NotHandledLine v1, NotHandledLine v2)
        {
            return (v1.CodeLine == v2.CodeLine &&
                    v1.FileName == v2.FileName &&
                    v1.Lines.Item1.StartInt == v2.Lines.Item1.StartInt &&
                    v1.Lines.Item1.EndInt == v2.Lines.Item1.EndInt &&
                    v1.Lines.Item2 == v2.Lines.Item2);
        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }
}
