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
    [Wrapped(typeof(VBProject))]
    public class VBProjectWrapping : VBProjectWrappingBase
    {
        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public VBProject VBProject { get; }
        public VBProjectWrapping(VBProject vbproject)
        {
            this.VBProject = vbproject;
        }

        public void SaveAs(string FileName)
        {
            VBProject.SaveAs(FileName);
        }

        public void MakeCompiledFile()
        {
            VBProject.MakeCompiledFile();
        }

        public Microsoft.Vbe.Interop.Application Application => VBProject.Application;
        public Microsoft.Vbe.Interop.Application Parent => VBProject.Parent;
        public string HelpFile { set { VBProject.HelpFile = value; } get { return VBProject.HelpFile; } }
        public int HelpContextID { set { VBProject.HelpContextID = value; } get { return VBProject.HelpContextID; } }
        public string Description { set { VBProject.Description = value; } get { return VBProject.Description; } }
        public vbext_VBAMode Mode => VBProject.Mode;
        public References References => VBProject.References;
        public string Name { set { VBProject.Name = value; } get { return VBProject.Name; } }
        public VBE VBE => VBProject.VBE;
        public VBProjects Collection => VBProject.Collection;
        public vbext_ProjectProtection Protection => VBProject.Protection;
        public bool Saved => VBProject.Saved;
        public VBComponents VBComponents => VBProject.VBComponents;
        public vbext_ProjectType Type => VBProject.Type;
        public string FileName => VBProject.FileName;
        public string BuildFileName { set { VBProject.BuildFileName = value; } get { return VBProject.BuildFileName; } }

    }
}
