using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class ApplicationWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
    }
}
