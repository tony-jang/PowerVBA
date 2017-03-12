using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Wrap
{
    public class WrappedAttribute : Attribute
    {
        public Type TargetType { get; set; }

        public WrappedAttribute(Type targetType)
        {
            this.TargetType = targetType;
        }
    }
}
