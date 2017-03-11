using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Parser.Data
{
    interface BaseData
    {
        string Name { get; set; }

        int Line { get; set; }

        string Type { get; set; }

        string File { get; set; }

        Accessor Accessor { get; set; }
    }
}
