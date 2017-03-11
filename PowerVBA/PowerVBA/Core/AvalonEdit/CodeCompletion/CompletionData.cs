using ICSharpCode.AvalonEdit.CodeCompletion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;
using System.Windows.Media;

namespace PowerVBA.Core.AvalonEdit.CodeCompletion
{
    class CompletionData : ICompletionData
    {
        protected object _Content;
        protected object _Description;
        protected ImageSource _Image;
        protected string _Text;
        protected string _RealConvText;


        public object Content { get { return _Content; } }
        public object Description { get { return _Description; } }
        public ImageSource Image { get { return _Image; } }
        public double Priority { get { return 1; } }
        public string Text { get { return _Text; } }
        public string RealConvText { get { return _RealConvText; } }

        public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
        {
            textArea.Document.Replace(completionSegment, RealConvText);
        }
        

        /// <param name="Text">Item1 : 인식용, Item2 : 실제 변환</param>
        /// <param name="Description">설명</param>
        /// <param name="Image">이미지</param>
        public CompletionData((string, string) Text, object Description = null, ImageSource Image = null)
        {
            _Text = Text.Item1;
            _RealConvText = Text.Item2;
            _Content = Text.Item1;
            _Description = Description;
            _Image = Image;
        }



        /// <param name="Text">인식 + 변환이 같습니다.</param>
        /// <param name="Description">설명</param>
        /// <param name="Image">이미지</param>
        public CompletionData(string Text, object Description = null, ImageSource Image = null)
        {
            _Text = Text;
            _RealConvText = Text;
            _Content = Text;
            _Description = Description;
            _Image = Image;
        }
    }

    public enum CompletionHotKey
    {
        Standard,
        DoubleTab
    }
}
