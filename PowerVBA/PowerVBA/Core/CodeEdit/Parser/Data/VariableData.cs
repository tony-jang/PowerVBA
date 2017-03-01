using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Parser.Data
{
    public class VariableData : BaseData
    {

        public string Name { get; set; }

        public int Line { get; set; }

        public string Type { get; set; }

        public string File { get; set; }

        public Accessor Accessor { get; set; }
    }
}
