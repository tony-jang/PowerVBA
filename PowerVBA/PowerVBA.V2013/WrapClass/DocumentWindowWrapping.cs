using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Wrap;

namespace PowerVBA.V2013.WrapClass

{
    [Wrapped(typeof(DocumentWindow))]
    class DocumentWindowWrapping : DocumentWindowWrappingBase
	{
        public override PPTVersion ClassVersion => PPTVersion.PPT2013;
        public DocumentWindow DocumentWindow { get; }
        public DocumentWindowWrapping(DocumentWindow documentwindow)
        {
            this.DocumentWindow = documentwindow;
        }

        public void FitToPage()
        {
            DocumentWindow.FitToPage();
        }

        public void Activate()
        {
            DocumentWindow.Activate();
        }

        public void LargeScroll(int Down = 1, int Up = 0, int ToRight = 0, int ToLeft = 0)
        {
            DocumentWindow.LargeScroll(Down, Up, ToRight, ToLeft);
        }

        public void SmallScroll(int Down = 1, int Up = 0, int ToRight = 0, int ToLeft = 0)
        {
            DocumentWindow.SmallScroll(Down, Up, ToRight, ToLeft);
        }

        public DocumentWindow NewWindow()
        {
            return DocumentWindow.NewWindow();
        }

        public void Close()
        {
            DocumentWindow.Close();
        }
        public dynamic RangeFromPoint(int X, int Y)
        {
            return DocumentWindow.RangeFromPoint(X, Y);
        }

        public int PointsToScreenPixelsX(float Points)
        {
            return DocumentWindow.PointsToScreenPixelsX(Points);
        }

        public int PointsToScreenPixelsY(float Points)
        {
            return DocumentWindow.PointsToScreenPixelsY(Points);
        }

        public void ScrollIntoView(float Left, float Top, float Width, float Height, MsoTriState Start = MsoTriState.msoTrue)
        {
            DocumentWindow.ScrollIntoView(Left, Top, Width, Height, Start);
        }

        public bool IsSectionExpanded(int sectionIndex)
        {
            return DocumentWindow.IsSectionExpanded(sectionIndex);
        }

        public void ExpandSection(int sectionIndex, bool Expand)
        {
            DocumentWindow.ExpandSection(sectionIndex, Expand);
        }

        public Microsoft.Office.Interop.PowerPoint.Application Application => DocumentWindow.Application;
        public dynamic Parent => DocumentWindow.Parent;
        public Selection Selection => DocumentWindow.Selection;
        public View View => DocumentWindow.View;
        public Presentation Presentation => DocumentWindow.Presentation;
        public PpViewType ViewType { set { DocumentWindow.ViewType = value; } get { return DocumentWindow.ViewType; } }
        public MsoTriState BlackAndWhite { set { DocumentWindow.BlackAndWhite = value; } get { return DocumentWindow.BlackAndWhite; } }
        public MsoTriState Active => DocumentWindow.Active;
        public PpWindowState WindowState { set { DocumentWindow.WindowState = value; } get { return DocumentWindow.WindowState; } }
        public string Caption => DocumentWindow.Caption;
        public float Left { set { DocumentWindow.Left = value; } get { return DocumentWindow.Left; } }
        public float Top { set { DocumentWindow.Top = value; } get { return DocumentWindow.Top; } }
        public float Width { set { DocumentWindow.Width = value; } get { return DocumentWindow.Width; } }
        public float Height { set { DocumentWindow.Height = value; } get { return DocumentWindow.Height; } }
        public int HWND => DocumentWindow.HWND;
        public Pane ActivePane => DocumentWindow.ActivePane;
        public Panes Panes => DocumentWindow.Panes;
        public int SplitVertical { set { DocumentWindow.SplitVertical = value; } get { return DocumentWindow.SplitVertical; } }
        public int SplitHorizontal { set { DocumentWindow.SplitHorizontal = value; } get { return DocumentWindow.SplitHorizontal; } }
    }
}
