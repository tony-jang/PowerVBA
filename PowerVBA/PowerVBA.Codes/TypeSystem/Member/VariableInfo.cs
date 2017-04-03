using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public class VariableInfo : IMember
    {
        public VariableInfo(string name)
        {
            Name = name;
        }
        public VariableInfo(string name, IType ReturnType)
        {
            Name = name;
            this.ReturnType = ReturnType;
        }
        public string Name { get; set; }
        public IType ReturnType { get; set; }
    }
}
