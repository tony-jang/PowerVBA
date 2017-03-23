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
            _StartInt = value;
            _EndInt = value;
        }

        public static implicit operator RangeInt(int i)
        {
            return new RangeInt(i);
        }

        private int _StartInt;
        public int StartInt
        {
            get { return _StartInt; }
            set { if (value <= _EndInt) _StartInt = value; }
        }
        private int _EndInt;
        public int EndInt
        {
            get { return _EndInt; }
            set { if (value <= _StartInt) _EndInt = value; }
        }

        public int[] GetRange()
        {
            List<int> intList = new List<int>();
            for (int i = StartInt; i <= EndInt; i++)
                intList.Add(i);
            return intList.ToArray();
        }
    }
}
