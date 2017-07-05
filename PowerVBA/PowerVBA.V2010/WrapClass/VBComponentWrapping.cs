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
    [Wrapped(typeof(_VBComponent))]
    public class VBComponentWrapping : VBComponentWrappingBase
    {
        public VBComponent VBComponent { get; }
        public VBComponentWrapping(VBComponent VBComponent) : base($"{VBComponent.Name}{GetExtensions(VBComponent.Type)}")
        {
            this.VBComponent = VBComponent;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public void Export(string FileName)
        {
            VBComponent.Export(FileName);
        }

        public Window DesignerWindow()
        {
            return VBComponent.DesignerWindow();
        }

        public void Activate()
        {
            VBComponent.Activate();
        }

        public bool Saved => VBComponent.Saved;
        public string Name { set { VBComponent.Name = value; OnNameChanged(); } get { return VBComponent.Name; } }
        public dynamic Designer => VBComponent.Designer;
        public CodeModule CodeModule => VBComponent.CodeModule;
        public vbext_ComponentType Type => VBComponent.Type;
        public VBE VBE => VBComponent.VBE;
        public VBComponents Collection => VBComponent.Collection;
        public bool HasOpenDesigner => VBComponent.HasOpenDesigner;
        public Properties Properties => VBComponent.Properties;
        public string DesignerID => VBComponent.DesignerID;

        public override string CompName => $"{Name}{GetExtensions(Type)}";

        public override string Code {
            get => (VBComponent.CodeModule.CountOfLines != 0 ? VBComponent.CodeModule.get_Lines(1, VBComponent.CodeModule.CountOfLines) : "");
            set { CodeModule.DeleteLines(1, CodeModule.CountOfLines);
                  CodeModule.AddFromString(value); }
        }

        public override int GetComponentType()
        {
            switch (Type)
            {
                case vbext_ComponentType.vbext_ct_Document:
                    return 1;
                case vbext_ComponentType.vbext_ct_StdModule:
                    return 2;
                case vbext_ComponentType.vbext_ct_MSForm:
                    return 3;
                case vbext_ComponentType.vbext_ct_ClassModule:
                    return 4;
                default:
                    return 0;
            }
        }

        public override string GetExtension => GetExtensions(Type);

        public static string GetExtensions(vbext_ComponentType Type)
        {
            switch (Type)
            {
                case vbext_ComponentType.vbext_ct_StdModule:
                    return ".bas";
                case vbext_ComponentType.vbext_ct_ClassModule:
                case vbext_ComponentType.vbext_ct_Document:
                    return ".cls";
                case vbext_ComponentType.vbext_ct_MSForm:
                    return ".frm";
                default:
                    return "";
            }
        }

        public override void SetCode(string code)
        {
            CodeModule.AddFromString(code);
        }

        public override void CodeClear()
        {
            CodeModule.DeleteLines(1, CodeModule.CountOfLines);
        }
    }
}
