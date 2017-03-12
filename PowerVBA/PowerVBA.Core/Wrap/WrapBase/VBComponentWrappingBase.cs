using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class VBComponentWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
    }
}
