using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeNodes
{
    class TextItem : CodeItemBase
    {
        public string Text { get; set; }

        public new string StrDescription { get => $"'{Text}'"; }
    }
}
