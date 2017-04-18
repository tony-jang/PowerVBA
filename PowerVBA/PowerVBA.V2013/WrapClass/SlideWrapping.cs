using core=Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;

namespace PowerVBA.V2013.WrapClass
{
    [Wrapped(typeof(Slide))]
    public class SlideWrapping : SlideWrappingBase
    {
        public Slide Slide { get; }
        public SlideWrapping(Slide slide)
        {
            this.Slide = slide;   
        }
        public override PPTVersion ClassVersion => PPTVersion.PPT2013;

        public Application Application => Slide.Application;
        public ShapeRange Background => Slide.Background;
        public core.MsoBackgroundStyleIndex BackgroundStyle { get { return Slide.BackgroundStyle; } set { Slide.BackgroundStyle = value; } }
        public ColorScheme ColorScheme { get { return Slide.ColorScheme; } set { Slide.ColorScheme = value; } }
        public Comments Comments => Slide.Comments;
        public CustomerData CustomerData => Slide.CustomerData;
        public CustomLayout CustomLayout { get { return Slide.CustomLayout; } set { Slide.CustomLayout = value; } }
        public Design Design { get { return Slide.Design; } set { Slide.Design = value; } }
        public core.MsoTriState DisplayMasterShapes { get { return Slide.DisplayMasterShapes; } set { Slide.DisplayMasterShapes = value; } }
        public core.MsoTriState FollowMasterBackground { get { return Slide.FollowMasterBackground; } set { Slide.FollowMasterBackground = value; } }
        public core.MsoTriState HasNotesPage => Slide.HasNotesPage;
        public HeadersFooters HeadersFooters => Slide.HeadersFooters;
        public Hyperlinks Hyperlinks => Slide.Hyperlinks;
        public PpSlideLayout Layout { get { return Slide.Layout; } set { Slide.Layout = value; } }
        public Master Master => Slide.Master;
        public string Name { get { return Slide.Name; } set { Slide.Name = value; } }
        public SlideRange NotesPage => Slide.NotesPage;
        public dynamic Parent => Slide.Parent;
        public int PrintSteps => Slide.PrintSteps;
        public core.Scripts Scripts => Slide.Scripts;
        public int sectionIndex => Slide.sectionIndex;
        public int SectionNumber => Slide.SectionNumber;
        public Shapes Shapes => Slide.Shapes;
        public int SlideID => Slide.SlideID;
        public int SlideIndex => Slide.SlideIndex;
        public int SlideNumber => Slide.SlideNumber;
        public SlideShowTransition SlideShowTransition => Slide.SlideShowTransition;
        public Tags Tags => Slide.Tags;
        public core.ThemeColorScheme ThemeColorScheme => Slide.ThemeColorScheme;
        public TimeLine TimeLine => Slide.TimeLine;
        public void ApplyTemplate(string FileName)
        {
            Slide.ApplyTemplate(FileName);
        }

        public void ApplyTemplate2(string FileName, string VariantGUID)
        {
            Slide.ApplyTemplate2(FileName, VariantGUID);
        }

        public void ApplyTheme(string themeName)
        {
            Slide.ApplyTheme(themeName);
        }

        public void ApplyThemeColorScheme(string themeColorSchemeName)
        {
            Slide.ApplyThemeColorScheme(themeColorSchemeName);
        }

        public void Copy()
        {
            Slide.Copy();
        }

        public void Cut()
        {
            Slide.Cut();
        }

        public void Delete()
        {
            Slide.Delete();
        }

        public SlideRange Duplicate()
        {
            return Slide.Duplicate();
        }

        public void Export(string FileName, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0)
        {
            Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight);
        }

        public void MoveTo(int toPos)
        {
            Slide.MoveTo(toPos);
        }

        public void MoveToSectionStart(int toSection)
        {
            Slide.MoveToSectionStart(toSection);
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false)
        {
            Slide.PublishSlides(SlideLibraryUrl, Overwrite, UseSlideOrder);
        }

        public void Select()
        {
            Slide.Select();
        }






    }
}
