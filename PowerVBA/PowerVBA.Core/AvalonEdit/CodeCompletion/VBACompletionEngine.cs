using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Global.RegexExpressions;
using PowerVBA.RegexPattern;
using PowerVBA.Core.Resources;
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
    public class VBACompletionEngine
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
            string prevS = "";

            if (dl.Length >= 2)
            {
                prevS = document.Text.Substring(dl.Offset, dl.Length - 1);
            }
            
            // 빈칸 또는 
            if (CodePattern.Default.IsMatch(s) || CodePattern.BlankNullPattern.IsMatch(s))
            {
                var CompletionDatas = 
                    new List<CompletionData>() { VBACompletions.Comp_Dim, VBACompletions.Comp_Public, VBACompletions.Comp_Private, VBACompletions.Comp_Do,
                        VBACompletions.Comp_For, VBACompletions.Comp_ForEach, VBACompletions.Comp_Select, VBACompletions.Comp_SelectCase, VBACompletions.Comp_End };
                data.AddRange(CompletionDatas);
            }

            if (CodePattern.ForBeforeTo.IsMatch(s) || CodePattern.ForBeforeTo_Ex.IsMatch(s)) { data.Add(VBACompletions.Comp_To); }
            if (CodePattern.ForBeforeStep.IsMatch(s) || CodePattern.ForBeforeStep_Ex.IsMatch(s)) { data.Add(VBACompletions.Comp_Step); }
            if (CodePattern.SCBeforeCase.IsMatch(s)) data.Add(VBACompletions.Comp_Case);
            if (CodePattern.ForBeforeAs_Ex.IsMatch(s) || CodePattern.VarBefore_As.IsMatch(s)) { data.Add(VBACompletions.Comp_As); }
            

            if (CodePattern.ForEach_BeforeEach.IsMatch(s)) { data.Add(VBACompletions.Comp_Each); }

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
