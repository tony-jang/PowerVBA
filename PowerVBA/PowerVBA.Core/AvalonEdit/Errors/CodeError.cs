using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.AvalonEdit.Errors
{
    public class CodeError
    {
        public CodeError(string ErrorCode, string Description, string FileName, int Line)
        {
            this.ErrorCode = ErrorCode;
            this.Description = Description;
            this.FileName = FileName;
            this.Line = Line;
        }
        public string ErrorCode { get; set; }
        public string Description { get; set; }
        public string FileName { get; set; }
        public int Line { get; set; }
    }
}
