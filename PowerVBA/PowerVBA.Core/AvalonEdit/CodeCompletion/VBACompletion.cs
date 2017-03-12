using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Core.AvalonEdit.CodeCompletion
{
    public class VBACompletion
    {
        public VBACompletion()
        {
        }

        protected VBACompletionEngine engine;

        public CodeCompletionResult GetCompletions(IDocument document, int offset)
        {
            return GetCompletions(document, offset,false);
        }

        public CodeCompletionResult GetCompletions(IDocument document, int offset, bool controlSpace)
        {
            return GetCompletions(document, offset, false, null);
        }

        public CodeCompletionResult GetCompletions(IDocument document, int offset, bool controlSpace, string variables)
        {
            engine = new VBACompletionEngine(document);

            var result = new CodeCompletionResult();

            char completionChar;

            try { completionChar = document.GetText(offset - 1, 1).ToCharArray()[0]; }
            catch (Exception)
            { completionChar = '\0'; }
            

            int startPos = offset;

            TextLocation tl = document.GetLocation(offset);
            


            Console.WriteLine("Getcompletions Call");
            

            if (char.IsLetterOrDigit(completionChar) || completionChar == '_')
            {
                if (startPos > 1 && char.IsLetterOrDigit(document.GetCharAt(startPos - 2)))
                    return result;

                result.TriggerWordLength = 1;
                result.TriggerWord = document.GetText(offset - 1, 1);

                result.CompletionData = engine.GetCompletionData(offset, false).ToList();
            }
            else
            {
                result.TriggerWordLength = 0;
                result.CompletionData = engine.GetCompletionData(offset, false).ToList();
            }
                
            
            return result;
        }
    }
}
