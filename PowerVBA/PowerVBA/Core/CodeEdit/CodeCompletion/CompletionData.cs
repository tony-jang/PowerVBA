using ICSharpCode.AvalonEdit.CodeCompletion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;
using System.Windows.Media;

namespace PowerVBA.Core.CodeEdit.CodeCompletion
{
    class CompletionData : ICompletionData
    {
        protected object _Content;
        protected object _Description;
        protected ImageSource _Image;
        protected string _Text;


        public object Content { get { return _Content; } }
        public object Description { get { return _Description; } }
        public ImageSource Image { get { return _Image; } }
        public double Priority { get { return 1; } }
        public string Text { get { return _Text; } }

        public CompletionFlag Flag { get; set; }


        public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
        {
            textArea.Document.Replace(completionSegment, Text);
        }

        public CompletionData(string Text, object Content, object Description = null, ImageSource Image = null, CompletionFlag flag = CompletionFlag.None)
        {
            _Text = Text;
            _Content = Content;
            _Description = Description;
            _Image = Image;
            Flag = flag;
        }
    }
}
