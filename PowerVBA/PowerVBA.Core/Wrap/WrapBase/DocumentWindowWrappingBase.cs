using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Core.Connector;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class DocumentWindowWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
    }
}
