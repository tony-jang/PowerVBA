using PowerVBA.Core.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Core.Interface;
using PowerVBA.V2013.Wrap.WrapClass;
using Microsoft.Office.Core;
using static PowerVBA.Core.Extension.boolEx;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Core.Wrap.WrapBase;
using Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;

namespace PowerVBA.V2013.Connector
{
    public class PPTConnector2013 : PPTConnectorBase
    {
        public PPTConnector2013()
        {
            Application = new ApplicationWrapping(new Application());
        }
        public PPTConnector2013(string FileLocation, bool NewFile = false, bool OpenWithWindow = true) : this()
        {
            Presentation = new PresentationWrapping(Application.Presentations.Open(FileLocation, MsoTriState.msoFalse, NewFile.BoolToState(), OpenWithWindow.BoolToState()));
            VBProject = new VBProjectWrapping(Presentation.VBProject);
        }
        public PPTConnector2013(bool OpenWithWindow = true) : this()
        {
            Presentation = new PresentationWrapping(Application.Presentations.Add(OpenWithWindow.BoolToState()));
            VBProject = new VBProjectWrapping(Presentation.VBProject);
            Presentation.Slides.AddSlide(1, Presentation.SlideMaster.CustomLayouts[1]);
        }




        public PresentationWrapping Presentation { get => (PresentationWrapping)_Presentation; set => _Presentation = value; }
        public ApplicationWrapping Application { get => (ApplicationWrapping)_PPTApp; set => _PPTApp = value; }    
        public VBProjectWrapping VBProject { get => (VBProjectWrapping)_VBProject; set => _VBProject = value; }

        public override PPTVersion Version => PPTVersion.PPT2013;

        public bool IsContainsName(string name)
        {
            foreach (VBA.VBComponent comp in VBProject.VBComponents)
            {
                if (name == comp.Name) return true;
            }

            return false;
        }

        #region [  Class/Form/Module 추가/제거  ]

        public override bool AddClass(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;
            if (IsContainsName(name)) return false;
            VBComponentWrapping newStandardModule = new VBComponentWrapping(VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule))
            {
                Name = name
            };
            return true;
        }

        public override bool AddForm(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;
            if (IsContainsName(name)) return false;
            VBComponentWrapping newStandardForm = new VBComponentWrapping(VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_MSForm))
            {
                Name = name
            };
            return true;
        }

        public override bool AddModule(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;
            if (IsContainsName(name)) return false;
            VBComponentWrapping newStandardModule = new VBComponentWrapping(VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule))
            {
                Name = name
            };
            return true;
        }
        
        public override bool DeleteClass(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;

            VBProject.VBComponents.Remove(VBProject.VBComponents.Cast<VBA.VBComponent>().Where((i) => (i.Name == name && i.Type == VBA.vbext_ComponentType.vbext_ct_ClassModule)).First());

            return true;
        }

        public override bool DeleteForm(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;

            VBProject.VBComponents.Remove(VBProject.VBComponents.Cast<VBA.VBComponent>().Where((i) => (i.Name == name && i.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)).First());
            return true;
        }

        public override bool DeleteModule(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Pattern.NamePattern)) return false;

            VBProject.VBComponents.Remove(VBProject.VBComponents.Cast<VBA.VBComponent>().Where((i) => (i.Name == name && i.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)).First());

            return true;
        }

        #endregion
        
        #region [  Slide 추가/제거  ]

        public override bool AddSlide()
        {
            int SlideNumber = 0;

            if (Presentation.Slides.Count != 0) SlideNumber = Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            try
            {
                Presentation.Slides.AddSlide(SlideNumber + 1, Presentation.Presentation.SlideMaster.CustomLayouts[1]);
                Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
            }
            catch (Exception) { return false; }
            

            return true;
        }

        public override bool AddSlide(int SlideNumber)
        {

            if (Presentation.Presentation.Slides.Count == 0) return false;
            
            try
            {
                Presentation.Slides.AddSlide(SlideNumber, Presentation.Presentation.SlideMaster.CustomLayouts[1]);
                Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
            }
            catch (Exception) { return false; }
            

            return true;
        }

        public override bool DeleteSlide()
        {
            try
            {
                Presentation.Slides[Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex].Delete();
            }
            catch (Exception) { return false; }
            

            return true;
        }

        public override bool DeleteSlide(int SlideNumber)
        {
            try
            {
                Presentation.Slides[SlideNumber].Delete();
            }
            catch (Exception) { return false; }


            return true;
        }

        #endregion



        public override List<ShapeWrappingBase> Shapes()
        {
            List<ShapeWrappingBase> shapes = new List<ShapeWrappingBase>();
            Presentation.Slides.Cast<Slide>()
                               .ToList()
                               .ForEach(i => 
                                   shapes.AddRange(i.Shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                                                           .ToList()
                                                           .Select((s) => new ShapeWrapping(s))));

            return shapes;
        }


        public override void Dispose()
        {
            Presentation.Close();
            if (Application.Presentations.Count == 0) Application.Quit();
        }

        
    }
}
