using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class TextItem : CodeItemBase
    {
        public TextItem(string Text, string FileName, string Name, (int, int) Segment) : base(FileName, Segment)
        {
            this.Text = Text;
        }
        public TextItem(string Text, string Name, (int,int) Segment) : base("", Segment)
        {
            this.Text = Text;
        }

        public string Text { get; set; }

        public new string StrDescription { get => $"'{Text}'"; }
    }
}
