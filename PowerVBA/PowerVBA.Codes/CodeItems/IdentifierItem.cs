using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    /// <summary>
    /// IdentifierCodeItem
    /// </summary>
    class IdentifierItem : CodeItemBase
    {
        public IdentifierItem(string Identifier, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            this.Identifier = Identifier;
        }
        public IdentifierItem(string Identifier, (int, int) Segment) : base(string.Empty, Segment)
        {
            this.Identifier = Identifier;
        }
        public string Identifier { get; set; }
        public new string StrDescription { get => $"'{Identifier}'"; }

    }
}
