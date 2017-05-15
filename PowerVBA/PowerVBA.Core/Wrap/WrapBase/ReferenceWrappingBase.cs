using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class ReferenceWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
        public abstract string Name { get; }
        public abstract string FullPath { get; }
    }
}
