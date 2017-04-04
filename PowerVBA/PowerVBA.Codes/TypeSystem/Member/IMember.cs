using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IMember
    {
        string Name { get; set; }
        Accessor Accessor { get; set; }
    }
}
