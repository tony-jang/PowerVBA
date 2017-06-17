using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Parsing
{
    public class Variable
    {
        public Variable(string name, int line)
        {
            this.Name = name;
            this.Line = line;
        }

        public string Namespace { get; set; }

        /// <summary>
        /// 변수의 이름입니다.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 타입 이름입니다.
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// 변수가 선언되어 있는 줄입니다.
        /// </summary>
        public int Line { get; set; }

        /// <summary>
        /// 사용하고 있는 줄입니다. 오프셋은 줄 기준으로의 오프셋입니다. (line | Offset)
        /// </summary>
        public List<(int, int)> Usages;
    }
}
