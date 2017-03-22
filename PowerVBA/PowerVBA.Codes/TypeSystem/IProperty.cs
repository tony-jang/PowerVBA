using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IUnresolvedProperty : IUnresolvedParameterizedMember
    {
        /// <summary>
        /// 
        /// </summary>
        bool IsLet { get; }
        bool IsGet { get; }

        new IProperty Resolve(ITypeResolveContext context);
    }

    public interface IProperty : IParameterizedMember
    {
        bool IsLet { get; }
        bool IsGet { get; }
    }
}
