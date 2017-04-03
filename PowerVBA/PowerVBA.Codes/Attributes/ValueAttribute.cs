using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Attributes
{
    class ValueAttribute : Attribute
    {
        public ValueAttribute(string Value)
        {
            this.Value = Value;
        }

        public string Value { get; set; }
    }
}
