using Microsoft.Vbe.Interop;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;

namespace PowerVBA.V2013.WrapClass

{
    [Wrapped(typeof(VBComponent))]
    public class VBComponentWrapping : VBComponentWrappingBase
    {
        public VBComponent VBComponent { get; }
        public VBComponentWrapping(VBComponent vbcomponent)
        {
            this.VBComponent = vbcomponent;
        }

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
        public override PPTVersion ClassVersion => PPTVersion.PPT2013;

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
        
        public override string CompName => Name;

    }
}
