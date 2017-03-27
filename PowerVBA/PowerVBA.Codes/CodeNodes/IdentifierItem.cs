using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeNodes
{
    /// <summary>
    /// IdentifierCodeItem
    /// </summary>
    class IdentifierItem : CodeItemBase
    {
        public new string Name { get => "Identifier"; }
        public new string StrDescription { get => $"'{Identifier}'"; }

    }
}
