using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Rendering;
using PowerVBA.Core.Error;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using static PowerVBA.Global.Globals;

namespace PowerVBA.Core.CodeEdit.Renderer
{
    class ErrorLineBackgroundRenderer : IBackgroundRenderer
    {
        public TextEditor Editor { get; }
        public List<CodeError> Errors { get; }

        public ErrorLineBackgroundRenderer(TextEditor editor, List<CodeError> Errors)
        {
            this.Editor = editor;
            this.Errors = Errors;
        }

        public KnownLayer Layer
        {
            get { return KnownLayer.Caret; }
        }

        public void Draw(TextView textView, DrawingContext drawingContext)
        {
            if (Editor.Document == null)
                return;

            textView.EnsureVisualLines();
            

            foreach (var err in Errors)
            {
                if (Editor.Document.LineCount < err.LineNumber) continue;
                var currentLine = Editor.Document.GetLineByNumber(err.LineNumber);
                string data = Editor.Text.Substring(currentLine.Offset, currentLine.Length).Replace("\t","    ");

                foreach (var rect in BackgroundGeometryBuilder.GetRectsForSegment(textView, currentLine))
                {
                    try
                    {
                        drawingContext.DrawRectangle(
                        null, new Pen(Brushes.Red, 0.5),
                        new Rect(new Point(rect.Location.X + 1, rect.Location.Y + rect.Height - 1), new Size(MeasureString(data, Editor).Width, 1)));
                    }
                    catch (Exception)
                    { }

                }
            }
            
        }
    }
}
