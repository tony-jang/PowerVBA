using PowerVBA.Codes.CodeItems;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    /// <summary>
    /// 현재 프로젝트 코드에 대한 모든 정보를 담고 있습니다.
    /// </summary>
    public class CodeInfo
    {
        public List<string> CustomMethods { get; set; }
        /// <summary>
        /// 현재 프로젝트 코드의 오류를 나타냅니다.
        /// </summary>
        public List<Error> ErrorList { get; set; }
        
        /// <summary>
        /// CodeInfo의 Node들 정보입니다.
        /// </summary>
        public List<CodeItemBase> Childrens { get; set; }

        public CodeInfo()
        {
            ErrorList = new List<Error>();
            Childrens = new List<CodeItemBase>();
        }

    }
}
