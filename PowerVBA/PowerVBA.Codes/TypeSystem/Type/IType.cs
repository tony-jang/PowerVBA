using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IType
    {
        string DisplayName { get; set; }
        string Namespace { get; set; }

        /// <summary>
        /// 유효한 타입인지에 대한 여부를 가져옵니다.
        /// </summary>
        bool IsVaild { get; set; }
        TypeKind TypeKind { get; set; }
    }

    public enum TypeKind
    {
        Class,
        Type,
    }
}
