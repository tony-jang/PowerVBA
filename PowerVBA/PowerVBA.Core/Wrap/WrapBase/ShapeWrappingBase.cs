using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;
using System.Windows.Media;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class ShapeWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }

        public abstract string ShapeType { get; }
        public abstract string Name { get; set; }
        public abstract float Width { get; set; }
        public abstract float Height { get; set; }
        public abstract float Left { get; set; }
        public abstract float Top { get; set; }
        public abstract Color RGB { get; }
        public abstract Color ForeRGB { get; }

        public abstract void Delete(out bool success);
    }
}
