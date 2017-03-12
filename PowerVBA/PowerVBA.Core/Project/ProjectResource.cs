using PowerVBA.Core.AvalonEdit.Parser.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Project
{
    class ProjectResource
    {
        public ProjectResource()
        {
            Variables = new List<VariableData>();
            Enums = new List<EnumData>();
            Lines = new List<ILineInfo>();
        }
        public List<VariableData> Variables;
        public List<EnumData> Enums;
        public List<ILineInfo> Lines;

    }
}
