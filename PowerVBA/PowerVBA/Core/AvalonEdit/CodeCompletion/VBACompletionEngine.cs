using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Global.RegexExpressions;
using PowerVBA.RegexPattern;
using PowerVBA.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace PowerVBA.Core.AvalonEdit.CodeCompletion
{
    class VBACompletionEngine
    {
        public class CompletionEngineCache
        {
            public CompletionEngineCache()
            {
                importCompletion = new List<ICompletionData>();
            }
            public List<ICompletionData> importCompletion;
        }

        protected CompletionEngineCache cache;
        protected IDocument document;
        protected int offset;

        /// <summary>
        /// End Of Line Marker
        /// </summary>
        protected string EolMarker = Environment.NewLine;
        protected string IndentStr = "\t";

        public VBACompletionEngine(IDocument document)
        {
            this.document = document;
            cache = new CompletionEngineCache();
            offset = 0;
        }

        public IEnumerable<ICompletionData> GetCompletionData(int offset, bool controlSpace)
        {
            var data = new List<CompletionData>();
            
            
            DocumentLine dl = (DocumentLine)document.GetLineByOffset(offset);
            string s = document.Text.Substring(dl.Offset, dl.Length);
            string prevS = document.Text.Substring(dl.Offset, dl.Length - 1);
            if (CodePattern.Default.IsMatch(s) || CodePattern.BlankNullPattern.IsMatch(s))
            {
                data.Add(VBACompletions.Comp_Dim);
                data.Add(VBACompletions.Comp_Public);
                data.Add(VBACompletions.Comp_Private);
                data.Add(VBACompletions.Comp_Do);
                data.Add(VBACompletions.Comp_For);
                data.Add(VBACompletions.Comp_ForEach);
                data.Add(VBACompletions.Comp_Select);
                data.Add(VBACompletions.Comp_SelectCase);
            }

            if (CodePattern.ForBeforeAs_Ex.IsMatch(s))
            {
                data.Add(VBACompletions.Comp_As);
            }

            //if (CodePattern.Second.IsMatch(s))
            //{
            //    data.Add(VBACompletions.Comp_As);
            //}


            return data;
        }

        public enum VisibleItem
        {
            None = 0,
            Enum = 1,
            Class = 2,
            Property = 4,
            Method = 8,
            Type = 16,
            Declarer = 32
        }
    }
}
