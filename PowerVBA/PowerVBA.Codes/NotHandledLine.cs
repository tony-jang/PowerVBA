using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    struct NotHandledLine
    {
        public NotHandledLine(string FileName, string CodeLine, RangeInt Lines)
        {
            this.FileName = FileName;
            this.CodeLine = CodeLine;
            this.Lines = Lines;
        }
        public string FileName { get; set; }
        public string CodeLine { get; set; }
        public RangeInt Lines { get; set; }


        public static NotHandledLine Empty = new NotHandledLine("", "", -1);


        public static bool operator !=(NotHandledLine v1, NotHandledLine v2)
        {
            return (v1.CodeLine != v2.CodeLine &&
                    v1.FileName != v2.FileName &&
                    v1.Lines.StartInt != v2.Lines.StartInt &&
                    v1.Lines.EndInt != v2.Lines.EndInt);
        }
        public static bool operator ==(NotHandledLine v1, NotHandledLine v2)
        {
            return (v1.CodeLine == v2.CodeLine &&
                    v1.FileName == v2.FileName &&
                    v1.Lines.StartInt == v2.Lines.StartInt &&
                    v1.Lines.EndInt == v2.Lines.EndInt);
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
