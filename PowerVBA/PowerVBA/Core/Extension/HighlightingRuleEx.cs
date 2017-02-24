using ICSharpCode.AvalonEdit.Highlighting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PowerVBA.Core.Extension
{
    static class HighlightingRuleEx
    {
        public static void Add(this HighlightingRule highlightrule, string data)
        {
            data = data.Replace(".", @"\.");


            string realData = Regex.Match(highlightrule.Regex.ToString(), @"\((.+)\)\\b").Groups[1].Value;

            List<string> strs = new List<string>();

            if (!realData.Contains("?>"))
            {
                strs = Regex.Split(realData, @"\|").ToList();
            }
            
            if (!strs.Contains(data)) strs.Add(data);

            strs.Sort();

            highlightrule.Regex = new Regex($@"\b({string.Join("|", strs)})\b", RegexOptions.IgnoreCase | 
                                                                                      RegexOptions.ExplicitCapture | 
                                                                                      RegexOptions.CultureInvariant);
        }
        public static void Remove(this HighlightingRule highlightrule, string data)
        {
            string realData = Regex.Match(highlightrule.Regex.ToString(), @"\((.+)\)\\b").Groups[1].Value;

            List<string> strs = new List<string>();

            if (!realData.Contains("?>"))
            {
                strs = Regex.Split(realData, @"\|").ToList();
            }

            if (strs.Contains(data)) strs.Remove(data);

            highlightrule.Regex = new Regex($@"\b({string.Join("|", strs)})\b", RegexOptions.IgnoreCase |
                                                                                      RegexOptions.ExplicitCapture |
                                                                                      RegexOptions.CultureInvariant);
        }
    }
}
