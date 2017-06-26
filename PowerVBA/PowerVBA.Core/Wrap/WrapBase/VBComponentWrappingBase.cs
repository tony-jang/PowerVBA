using PowerVBA.Codes;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class VBComponentWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
        public abstract string CompName { get; }
        public abstract string Code { get; set; }
        public CodeFile File { get; set; }

        public event BlankDelegate NameChanged;

        public VBComponentWrappingBase(string Name)
        {
            this.NameChanged += VBComponentWrappingBase_NameChanged;
            File = new CodeFile(Name);
        }

        public void OnNameChanged()
        {
            NameChanged?.Invoke();
        }

        private void VBComponentWrappingBase_NameChanged()
        {
            File.FileName = CompName;
        }
    }
}