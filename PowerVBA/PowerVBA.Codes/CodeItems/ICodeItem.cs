using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeItems
{
    public interface ICodeItem
    {
        
        (int,int) Segment { get; set; }

        /// <summary>
        /// 현재 CodeItem이 포함되어 있는 파일 이름입니다.
        /// </summary>
        string FileName { get; set; }
        string StrDescription { get; }
    }
}
