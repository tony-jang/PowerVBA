using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Editing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.Substitution
{
    interface ISubstitution
    {
        TextArea TextArea { get; }

        bool Handled { get; set; }

        /// <summary>
        /// TextEditor의 텍스트를 변경합니다.
        /// </summary>
        void Convert();


        int Indentation { get; set; }

    }
}
