using Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;
using System;

namespace PowerVBA.V2010.WrapClass

{
    [Wrapped(typeof(_Slide))]
    public class SlideWrapping : SlideWrappingBase
    {
        public Slide Slide { get; }
        public SlideWrapping(Slide Slide)
        {
            this.Slide = Slide;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public void Select()
        {
            Slide.Select();
        }

        public void Cut()
        {
            Slide.Cut();
        }

        public void Copy()
        {
            Slide.Copy();
        }

        public SlideRange Duplicate()
        {
            return Slide.Duplicate();
        }

        public void Delete()
        {
            Slide.Delete();
        }

        public void Export(string FileName, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0)
        {
            Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight);
        }

        public void MoveTo(int toPos)
        {
            Slide.MoveTo(toPos);
        }

        public void ApplyTemplate(string FileName)
        {
            Slide.ApplyTemplate(FileName);
        }

        public void ApplyTheme(string themeName)
        {
            Slide.ApplyTheme(themeName);
        }

        public void ApplyThemeColorScheme(string themeColorSchemeName)
        {
            Slide.ApplyThemeColorScheme(themeColorSchemeName);
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false)
        {
            Slide.PublishSlides(SlideLibraryUrl, Overwrite, UseSlideOrder);
        }

        public void MoveToSectionStart(int toSection)
        {
            Slide.MoveToSectionStart(toSection);
        }

        public Microsoft.Office.Interop.PowerPoint.Application Application => Slide.Application;
        public dynamic Parent => Slide.Parent;
        public Microsoft.Office.Interop.PowerPoint.Shapes Shapes => Slide.Shapes;
        public HeadersFooters HeadersFooters => Slide.HeadersFooters;
        public SlideShowTransition SlideShowTransition => Slide.SlideShowTransition;
        public ColorScheme ColorScheme { set { Slide.ColorScheme = value; } get { return Slide.ColorScheme; } }
        public Microsoft.Office.Interop.PowerPoint.ShapeRange Background => Slide.Background;
        public string Name { set { Slide.Name = value; } get { return Slide.Name; } }
        public int SlideID => Slide.SlideID;
        public int PrintSteps => Slide.PrintSteps;
        public PpSlideLayout Layout { set { Slide.Layout = value; } get { return Slide.Layout; } }
        public Tags Tags => Slide.Tags;
        public int SlideIndex => Slide.SlideIndex;
        public int SlideNumber => Slide.SlideNumber;
        public MsoTriState DisplayMasterShapes { set { Slide.DisplayMasterShapes = value; } get { return Slide.DisplayMasterShapes; } }
        public MsoTriState FollowMasterBackground { set { Slide.FollowMasterBackground = value; } get { return Slide.FollowMasterBackground; } }
        public SlideRange NotesPage => Slide.NotesPage;
        public Master Master => Slide.Master;
        public Hyperlinks Hyperlinks => Slide.Hyperlinks;
        public Scripts Scripts => Slide.Scripts;
        public Comments Comments => Slide.Comments;
        public Design Design { set { Slide.Design = value; } get { return Slide.Design; } }
        public TimeLine TimeLine => Slide.TimeLine;
        public int SectionNumber => Slide.SectionNumber;
        public CustomLayout CustomLayout { set { Slide.CustomLayout = value; } get { return Slide.CustomLayout; } }
        public ThemeColorScheme ThemeColorScheme => Slide.ThemeColorScheme;
        public MsoBackgroundStyleIndex BackgroundStyle { set { Slide.BackgroundStyle = value; } get { return Slide.BackgroundStyle; } }
        public CustomerData CustomerData => Slide.CustomerData;
        public int sectionIndex => Slide.sectionIndex;
        public MsoTriState HasNotesPage => Slide.HasNotesPage;

        
    }
}
