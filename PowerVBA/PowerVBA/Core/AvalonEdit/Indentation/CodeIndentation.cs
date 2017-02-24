using ICSharpCode.AvalonEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Core.AvalonEdit.Indentation
{
    class CodeIndentation
    {
        TextEditor Editor;
        public CodeIndentation(TextEditor editor)
        {
            this.Editor = editor;
        }
        public bool Handled = false;
        public void Indent()
        {
            if (Editor.SelectionLength != 0)
                return;
            if (Handled) return;

            int Offset = Editor.CaretOffset;

            List<string> sb = new List<string>();
            foreach(var line in Editor.Document.Lines)
            {
                string Data = Editor.Text.Substring(line.Offset, line.Length);

                sb.Add(Data);
            }
            Handled = true;
            string str = string.Join("\r\n", sb);
            Editor.Text = str;

            Editor.CaretOffset = Offset;

            Handled = false;

        }
    }
}
