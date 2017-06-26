using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Controls.Tools
{
    public class LinePointEventArgs : EventArgs
    {
        public LinePointEventArgs(LinePoint linePoint, string fileName)
        {
            this.LinePoint = linePoint;
            this.FileName = fileName;
        }
        public LinePoint LinePoint { get; internal set; }
        public string FileName { get; internal set; }
    }
}
