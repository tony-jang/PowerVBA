using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Comparer
{
    static class TextComparer
    {
        public static List<TextSegment> Compare(string before, string after)
        {
            if (before == after) return new List<TextSegment>();

            var list = new List<TextSegment>();

            string[] beforeLine = before.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            string[] afterLine = after.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            int beforeCtr = 0, afterCtr = 0;

            for (int i = 0; i< Math.Max(beforeLine.Count(), afterLine.Count()); i++)
            {

            }

            var seg = new TextSegment();

            return list;
        }
    }
}
