using ICSharpCode.AvalonEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Interface
{
    interface ISubstitution
    {
        TextEditor Editor { get; }

        bool Handled { get; set; }

        /// <summary>
        /// TextEditor의 텍스트를 변경합니다.
        /// </summary>
        void Convert();


        int Indentation { get; set; }

    }
}
