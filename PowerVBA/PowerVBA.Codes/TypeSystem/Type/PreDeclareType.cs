using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    class PreDeclareType : IType
    {
        public string DisplayName { get; set; }
        public string Namespace { get; set; }
        public bool IsVaild { get; set; }
        public TypeKind TypeKind { get; set; }
    }
}
