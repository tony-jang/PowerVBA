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
        /// <summary>
        /// 커스텀 메소드들을 가져옵니다.
        /// </summary>
        public List<MethodInfo> CustomMethods { get; }


        /// <summary>
        /// 현재 프로젝트에서 선언된 속성들을 가져옵니다.
        /// </summary>
        public PropertyInfo[] Properties
        {
            get => CustomMethods.Where((m) => m.GetType() == typeof(PropertyInfo)).Cast<PropertyInfo>().ToArray();
        }
        

        /// <summary>
        /// 현재 프로젝트에서 선언된 메소드중 반환 값이 없는 메소드만 가져옵니다.
        /// </summary>
        public BasisMethodInfo[] SubMethods
        {
            get => CustomMethods.Where((m) => m.GetType() == typeof(BasisMethodInfo)).Cast<BasisMethodInfo>().ToArray();
        }

        /// <summary>
        /// 현재 프로젝트에서 선언된 메소드중 반환하는 메소드만 가져옵니다.
        /// </summary>
        public FunctionInfo[] FuncMethods
        {
            get => CustomMethods.Where((m) => m.GetType() == typeof(FunctionInfo)).Cast<FunctionInfo>().ToArray();
        }
        

        //public VariableInfo FindVariable()
        //{

        //}
        //public MethodInfo FindMethod(int FindLine, string stringFindName)
        //{
            
        //}
        //public IMember FindMember()
        //{

        //}


        /// <summary>
        /// 현재 프로젝트의 변수들을 나타냅니다.
        /// </summary>
        public List<VariableInfo> Variables { get; set; }

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
            CustomMethods = new List<MethodInfo>();
            Variables = new List<VariableInfo>();
            
            ErrorList = new List<Error>();
            Childrens = new List<CodeItemBase>();
        }

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
