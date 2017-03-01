using PowerVBA.Core.CodeEdit.Parser.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Parser
{
    class ProjectResource
    {
        public ProjectResource()
        {
            Variables = new List<VariableData>();
            Enums = new List<EnumData>();
        }
        public List<VariableData> Variables;
        public List<EnumData> Enums;

    }
}
