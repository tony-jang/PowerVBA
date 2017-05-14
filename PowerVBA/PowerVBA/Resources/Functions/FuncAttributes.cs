using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Resources.Functions
{
    public enum FuncAttributes
    {
        [Description("Return")]
        ReturnType,
        [Description("Description")]
        Description,
        [Description("Dependency")]
        Dependency,
    }
}
