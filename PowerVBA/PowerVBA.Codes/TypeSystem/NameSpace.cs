using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public class NameSpace
    {
        public NameSpace(string[] NameSpaceString)
        {
            NameSpaces = NameSpaceString.ToList();
        }

        List<string> NameSpaces;

        public string DisplayName {
            get
            {
                return string.Join(".", NameSpaces);
            }
        }
    }
}
