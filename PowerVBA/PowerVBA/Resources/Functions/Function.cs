using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Resources.Functions
{
    public class Function
    {
        public Function(FuncName name, string code, bool isReturn, string returnType, string dependencyMessage, string description)
        {
            this.Name = name;
            this.Code = code;
            this.IsReturn = isReturn;
            this.ReturnType = returnType;
            this.DependencyMessage = dependencyMessage;
            this.Description = description;
        }
        /// <summary>
        /// 함수 이름입니다.
        /// </summary>
        public FuncName Name { get; set; }

        /// <summary>
        /// 코드 부분입니다.
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 코드 설명입니다.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 값을 반환하는지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsReturn { get; set; }

        /// <summary>
        /// 반환하는 타입을 가져옵니다. IsReturn이 False일 경우 없습니다.
        /// </summary>
        public string ReturnType { get; set; }

        /// <summary>
        /// 종속성 정보에 대한 메세지를 가져옵니다.
        /// </summary>
        public string DependencyMessage { get; set; }

        /// <summary>
        /// 사용되는 함수인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsUse { get; set; }
    }
}
