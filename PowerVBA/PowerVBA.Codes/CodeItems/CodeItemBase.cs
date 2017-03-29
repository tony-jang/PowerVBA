using System;
using ICSharpCode.AvalonEdit.Document;
using System.Collections.Generic;

namespace PowerVBA.Codes.CodeItems
{
    public abstract class CodeItemBase : ICodeItem
    {
        public CodeItemBase(string FileName, (int,int) Segment)
        {
            this.FileName = FileName;
            this.Segment = Segment;
            Childrens = new List<CodeItemBase>();
        }
        public (int,int) Segment { get; set; }
        public string FileName { get; set; }
        public string StrDescription { get; }
        public List<CodeItemBase> Childrens { get; set; }
        public override string ToString()
        {
            return StrDescription;
        }


        public CodeItemBase()
        {
            Childrens = new List<CodeItemBase>();
        }
    }
}