using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    class ParameterItem : CodeItemBase

    {
        public ParameterItem(ParamAccessor Accessor, string Name, string Type, string InitValue, string FileName, (int, int) Segment) : base(FileName, Segment)
        {
            this.Accessor = Accessor;
            this.Name = Name;
            this.Type = Type;
            this.InitValue = InitValue;
        }



        public ParamAccessor Accessor { get; set; }

        public string Name { get; set; }
        public string Type { get; set; }
        public string InitValue { get; set; }

    }
}
