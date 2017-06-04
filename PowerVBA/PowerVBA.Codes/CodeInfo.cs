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

        public CodeInfo()
        {
            ErrorList = new List<Error>();
            Childrens = new List<CodeItemBase>();
            Lines = new List<LineInfo>();

            Properties = new List<string>();
            Subs = new List<string>();
            Functions = new List<string>();
        }
        
        /// <summary>
        /// 현재 프로젝트 코드의 오류를 나타냅니다.
        /// </summary>
        public List<Error> ErrorList { get; set; }
        
        /// <summary>
        /// CodeInfo의 Node들 정보입니다.
        /// </summary>
        public List<CodeItemBase> Childrens { get; set; }
        public List<LineInfo> Lines { get; set; }


        /// <summary>
        /// 함수 목록입니다.
        /// </summary>
        public List<string> Functions { get; set; }
        /// <summary>
        /// 서브루틴 목록입니다.
        /// </summary>
        public List<string> Subs { get; set; }
        /// <summary>
        /// 프로퍼티 목록입니다. get/let과 set이 중복되도 하나로 처리합니다.
        /// </summary>
        public List<string> Properties { get; set; }
        

        public void Reset()
        {
            ErrorList = new List<Error>();
            Childrens = new List<CodeItemBase>();
        }

        public override string ToString()
        {
            List<string> returnD = new List<string>();

            foreach (var itm in Childrens)
            {
                returnD.Add(itm.ToString());
            }

            return string.Join("\r\n", returnD.ToArray());
        }
    }
}
