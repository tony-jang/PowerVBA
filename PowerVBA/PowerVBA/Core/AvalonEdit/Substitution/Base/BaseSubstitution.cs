using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit;

namespace PowerVBA.Core.AvalonEdit.Substitution.Base
{
    abstract class BaseSubstitution : ISubstitution
    {
        public BaseSubstitution(TextEditor editor)
        {
            this.Editor = editor;
        }

        public TextEditor Editor { get; set; }

        public bool Handled { get; set; } = false;

        public int Indentation { get; set; } = 0;

        public abstract void Convert();

        internal string GetIndentation()
        {
            StringBuilder sb = new StringBuilder();
            for(int i = 0; i < Indentation; i++)
            {
                sb.Append("\t");
            }

            return sb.ToString();
        }
    }
}
