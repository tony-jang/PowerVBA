using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public class PropertyInfo : MethodInfo
    {
        public PropertyInfo(string Name, PropertySetType setType) : base(Name)
        {
            SetType = setType;
        }

        public PropertySetType SetType { get; }
    }

    public enum PropertySetType
    {
        Get,
        Set,
        Let
    }
}
