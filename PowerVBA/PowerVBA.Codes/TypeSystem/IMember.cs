using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public interface IMember
    {
        string Name { get; set; }
        string FileName { get; set; }
        LinePoint LinePoint { get; set; }
        List<LinePoint> Usages { get; set; }
        bool IsPrivate { get; set; }
    }
}
