using PowerVBA.Codes.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeNodes
{

    
    class TypeItem : CodeItemBase
    {
        public ClassType ClassType { get; }

        public Accessor Accessor { get; }

        public new string StrDescription { get => $"{Accessor.ToString()} {ClassType.ToString()} {Name}"; }

    }
}
