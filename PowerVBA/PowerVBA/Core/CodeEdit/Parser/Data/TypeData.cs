using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Parser.Data
{
    public class TypeData : BaseData
    {
        public TypeData(string name, int line, string type, string file, Accessor Accessor)
        {
            Name = name;
            Line = line;
            Type = type;
            File = file;
            this.Accessor = Accessor;
            Variables = new List<VariableData>();
        }
        public Accessor Accessor { get; set; }
        public string File { get; set; }

        public int Line { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public List<VariableData> Variables;
    }
}
