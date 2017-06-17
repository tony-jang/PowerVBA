using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Parsing
{
    public class VariableManager : List<Variable>
    {
        public VariableManager() : base()
        {
        }

        public VariableManager(CodeInfo codeInfo, int capacity) : base(capacity)
        {
            this.CodeInfo = codeInfo;
        }

        public VariableManager(CodeInfo codeInfo, IEnumerable<Variable> collection) : base(collection)
        {
            this.CodeInfo = codeInfo;
        }

        CodeInfo CodeInfo { get; set; }

        public Variable FindName(string name)
        {
            return this.Where(i => i.Name.ToLower() == name.ToLower()).FirstOrDefault();
        }

        public bool IsExists(string name, LinePoint linePoint)
        {
            return this.Where(i => i.Name.ToLower() == name.ToLower()).Count() != 0;
        }
        
    }
}
