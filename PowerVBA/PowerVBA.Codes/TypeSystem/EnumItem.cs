using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Codes.Parsing;

namespace PowerVBA.Codes
{
    public class EnumItem : IMember
    {
        public EnumItem(string name, string fileName, int line, int offset)
        {
            this.Name = name;
            this.FileName = fileName;
            this.LinePoint = new LinePoint(line, name.Length, offset);

            this.Usages = new List<LinePoint>();
        }
        public string Name { get; set; }
        public string FileName { get; set; }
        public LinePoint LinePoint { get; set; }
        public List<LinePoint> Usages { get; set; }
        public bool IsPrivate { get; set; }
    }
}
