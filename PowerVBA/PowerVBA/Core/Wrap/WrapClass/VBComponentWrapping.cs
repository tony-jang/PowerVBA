using Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections;


namespace PowerVBA.Core.Wrap.WrapClass

{
    [Wrapped(typeof(VBComponent))]
    public class VBComponentWrapping : IWrappingClass
    {
        public VBComponent VBComponent { get; }
        public VBComponentWrapping(VBComponent vbcomponent)
        {
            this.VBComponent = vbcomponent;
        }

        public CodeModule CodeModule => VBComponent.CodeModule;
        public VBComponents Collection => VBComponent.Collection;
        public dynamic Designer => VBComponent.Designer;
        public string DesignerID => VBComponent.DesignerID;
        public bool HasOpenDesigner => VBComponent.HasOpenDesigner;
        public string Name { set { VBComponent.Name = value; } get { return VBComponent.Name; } }
        public Microsoft.Vbe.Interop.Properties Properties => VBComponent.Properties;
        public bool Saved => VBComponent.Saved;
        public vbext_ComponentType Type => VBComponent.Type;
        public VBE VBE => VBComponent.VBE;
        public void Activate()
        {
            VBComponent.Activate();
        }

        public Window DesignerWindow()
        {
            return VBComponent.DesignerWindow();
        }

        public void Export(string FileName)
        {
            VBComponent.Export(FileName);
        }

    }
}
