using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace PowerVBA.Core.Error
{
    class CodeError
    {

        public CodeError(int LineNum, ErrorType Type, string Message, List<Run> VisibleData = null)
        {
            LineNumber = LineNum;
            errorType = Type;
            ErrorMessage = Message;
            if (VisibleData == null) this.VisibleData = new List<Run>();
            else this.VisibleData = VisibleData;
        }

        public int LineNumber { get; set; }

        public ErrorType errorType { get; set; }

        public string ErrorMessage { get; set; }

        public List<Run> VisibleData { get; set; }


        #region [  미리 정의된 코드 오류  ]

        public static CodeError GetAccessorError(string wrongAccessor, int LineNum, string Message)
        {
            CodeError ce = new CodeError(LineNum, ErrorType.Error, Message);

            


            return ce;
        }

        #endregion

    }
}
