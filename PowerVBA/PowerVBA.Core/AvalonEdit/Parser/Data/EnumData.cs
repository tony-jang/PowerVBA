using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Parser.Data
{
    public class EnumData : BaseData
    {
        public EnumData(string name, int line, string type,string file, Accessor Accessor)
        {
            Name = name;
            Line = line;
            Type = type;
            File = file;
            this.Accessor = Accessor;
            Data = new List<SubEnum>();
        }
        public string Name { get; set; }

        public int Line { get; set; }

        public string Type { get; set; }

        public string File { get; set; }

        public List<SubEnum> Data { get; set; }

        public Accessor Accessor { get; set; }
    }
    public struct SubEnum
    {
        public string EnumData { get; set; }
        public int Index { get; set; }
    }
}
