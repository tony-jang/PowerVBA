using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public struct RangeInt
    {
        public RangeInt(int value)
        {
            StartInt = value;
            EndInt = value;
        }

        public static implicit operator RangeInt(int i)
        {
            return new RangeInt(i);
        }
        public int StartInt { get; set; }
        public int EndInt { get; set; }

        public int[] GetRange()
        {
            List<int> intList = new List<int>();
            for (int i = StartInt; i <= EndInt; i++)
                intList.Add(i);
            return intList.ToArray();
        }
    }
}
