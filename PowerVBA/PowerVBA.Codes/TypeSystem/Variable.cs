using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public class Variable : IMember
    {
        public Variable(string name, string fileName, int line, int offset)
        {
            this.Name = name;
            this.FileName = fileName;
            this.LinePoint = new LinePoint(line, name.Length, offset);

            this.Usages = new List<LinePoint>();
        }

        /// <summary>
        /// 변수의 이름입니다.
        /// </summary>
        public string Name { get; set; }

        public string FileName { get; set; }

        /// <summary>
        /// 타입 이름입니다.
        /// </summary>
        public string Type { get; set; }

        public LinePoint LinePoint { get; set; }

        public bool IsPrivate { get; set; }


        /// <summary>
        /// 사용하고 있는 줄입니다. 오프셋은 줄 기준으로의 오프셋입니다. (line | Offset)
        /// </summary>
        public List<LinePoint> Usages { get; set; }
    }
}
