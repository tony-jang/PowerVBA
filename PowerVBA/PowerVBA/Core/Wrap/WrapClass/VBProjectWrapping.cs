using Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Wrap.WrapClass
{
    [Wrapped(typeof(VBProject))]
    public class VBProjectWrapping : IWrappingClass
    {
        public VBProject VBProject { get; }

        public VBProjectWrapping(VBProject VBProject)
        {
            this.VBProject = VBProject;
        }

        public Application Application => VBProject.Application;
        public string BuildFileName { set { VBProject.BuildFileName = value; } get { return VBProject.BuildFileName; } }
        public VBProjects Collection => VBProject.Collection;
        public string Description { set { VBProject.Description = value; } get { return VBProject.Description; } }
        public string FileName => VBProject.FileName;
        public int HelpContextID { set { VBProject.HelpContextID = value; } get { return VBProject.HelpContextID; } }
        public string HelpFile { set { VBProject.HelpFile = value; } get { return VBProject.HelpFile; } }
        public vbext_VBAMode Mode => VBProject.Mode;
        public string Name { set { VBProject.Name = value; } get { return VBProject.Name; } }
        public Application Parent => VBProject.Parent;
        public vbext_ProjectProtection Protection => VBProject.Protection;
        public References References => VBProject.References;
        public bool Saved => VBProject.Saved;
        public vbext_ProjectType Type => VBProject.Type;
        public VBComponents VBComponents => VBProject.VBComponents;
        public VBE VBE => VBProject.VBE;
        public void MakeCompiledFile()
        {
            VBProject.MakeCompiledFile();
        }

        public void SaveAs(string FileName)
        {
            VBProject.SaveAs(FileName);
        }
    }
}
