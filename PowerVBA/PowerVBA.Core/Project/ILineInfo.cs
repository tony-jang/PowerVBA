using PowerVBA.Core.AvalonEdit.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Project
{
    public interface ILineInfo
    {

        int Offset { get; set; }
        int Length { get; set; }
        LineType LineType { get; set; }


        #region [  플래그  ]

        /// <summary>
        /// If문 안에 포함되어 있는지에 대한 여부를 나타냅니다.
        /// </summary>
        bool IsInIf { get; set; }
        /// <summary>
        /// Select문 안에 포함되어 있는지에 대한 여부를 나타냅니다.
        /// </summary>
        bool IsInSelect { get; set; }
        /// <summary>
        /// Enum문 안에 포함되어 있는지에 대한 여부를 나타냅니다.
        /// </summary>
        bool IsInEnum { get; set; }
        /// <summary>
        /// Type문 안에 포함되어 있는지에 대한 여부를 나타냅니다.
        /// </summary>
        bool IsInType { get; set; }
        /// <summary>
        /// Class문 안에 포함되어 있는지에 대한 여부를 나타냅니다.
        /// </summary>
        bool IsInClass { get; set; }
        



        #endregion
    }
}
