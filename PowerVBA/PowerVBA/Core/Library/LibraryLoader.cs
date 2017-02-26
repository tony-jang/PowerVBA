using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Library
{
    /// <summary>
    /// 동적으로 라이브러리를 가져옵니다.
    /// </summary>
    class LibraryLoader
    {
        public LibraryLoader(byte[] Data)
        {
            asm = Assembly.Load(Data);
        }

        Assembly asm;

    }
}
