using PowerVBA.Core.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Interface
{
    public interface IWrappingClass
    {
        PPTVersion ClassVersion { get; }
    }
}
