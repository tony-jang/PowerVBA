using System;
using ICSharpCode.AvalonEdit.Document;
using System.Collections.Generic;

namespace PowerVBA.Codes.CodeNodes
{
    public abstract class CodeItemBase : ICodeItem
    {
        public string Identifier { get; set; }
        public TextSegment segment { get; set; }
        public string FileName { get; set; }
        public string Name { get; }
        public string StrDescription { get; }
        public List<CodeItemBase> Childrens { get; set; }
        public override string ToString()
        {
            return $"{Name} : {StrDescription}";
        }


        public CodeItemBase()
        {
            Childrens = new List<CodeItemBase>();
        }
    }
}