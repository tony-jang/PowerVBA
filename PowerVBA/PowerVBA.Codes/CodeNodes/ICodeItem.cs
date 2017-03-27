using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.CodeNodes
{
    public interface ICodeItem
    {
        /// <summary>
        /// 식별자입니다.
        /// </summary>
        string Identifier { get; set; }
        
        TextSegment segment { get; set; }

        /// <summary>
        /// 현재 CodeItem이 포함되어 있는 파일 이름입니다.
        /// </summary>
        string FileName { get; set; }

        string Name { get; }
        string StrDescription { get; }
    }
}
