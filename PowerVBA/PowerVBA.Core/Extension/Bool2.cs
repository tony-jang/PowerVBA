using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Extension
{
    public struct Bool2
    {
        public static Bool2 True = new Bool2(true);
        public static Bool2 False = new Bool2(false);

        public bool Value { get; set; }

        public Bool2(bool value)
        {
            this.Value = value;
        }

        public static implicit operator MsoTriState(Bool2 value)
        {
            return value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }

        public static implicit operator Bool2(MsoTriState state)
        {
            return state == MsoTriState.msoTrue;
        }

        public static implicit operator Bool2(bool value)
        {
            return new Bool2(value);
        }

        public static implicit operator bool(Bool2 bool2)
        {
            return bool2.Value;
        }
    }
}
