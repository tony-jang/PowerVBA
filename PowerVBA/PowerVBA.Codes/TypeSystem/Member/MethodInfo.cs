using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Codes.Enums;

namespace PowerVBA.Codes.TypeSystem
{
    public abstract class MethodInfo : IMember
    {
        List<Parameter> Parameters { get; }
        public string Name { get; set; }
        public Accessor Accessor { get; set; }

        public MethodInfo(string Name, Accessor Accessor = Accessor.Dim)
        {
            this.Name = Name;
            Parameters = new List<Parameter>();
        }
    }
}
