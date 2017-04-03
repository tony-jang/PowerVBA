using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Attributes
{
    class PriorityAttribute : Attribute
    {
        public PriorityAttribute(int Priority)
        {
            this.Priority = Priority;
        }

        public int Priority { get; set; }
    }
}
