using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Error
{
    class CodeError
    {

        public CodeError(int LineNum, ErrorType Type, string Message)
        {
            LineNumber = LineNum;
            errorType = Type;
            ErrorMessage = Message;
        }

        public int LineNumber { get; set; }

        public ErrorType errorType { get; set; }

        public string ErrorMessage { get; set; }

    }
}
