using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    class NotHandledLine
    {
        public NotHandledLine(string FileName, string CodeLine, (RangeInt,int) Lines) 
        {
            this.FileName = FileName;
            this.CodeLine = CodeLine;
            this.Lines = Lines;
        }
        public string FileName { get; set; }
        public string CodeLine { get; set; }
        public (RangeInt, int) Lines { get; set; }
    }
}
