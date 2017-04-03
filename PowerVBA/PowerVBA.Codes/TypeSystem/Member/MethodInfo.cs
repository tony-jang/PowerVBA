using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public abstract class MethodInfo : IMember
    {
        List<Parameter> Parameters { get; }
        public string Name { get; set; }
        public IType ReturnType { get; set; }

        public MethodInfo(string Name)
        {
            this.Name = Name;
            Parameters = new List<Parameter>();
        }
    }
}
