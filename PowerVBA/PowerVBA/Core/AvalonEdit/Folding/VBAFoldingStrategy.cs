﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Document;
using System.Text.RegularExpressions;
using System.Windows;

namespace PowerVBA.Core.AvalonEdit.Folding
{
    class VBAFoldingStrategy
    {
        private string[] foldablewords = {"sub", "function" , "type", "enum", "class" };

        
        public IEnumerable<NewFolding> CreateNewFoldings(TextDocument document, out int firstErrorOffset)
        {
            firstErrorOffset = -1;

            var foldings = new List<NewFolding>();

            string text = document.Text;
            string lowerCaseText = text.ToLower();

            foreach (string foldableKeyword in foldablewords)
            {
                foldings.AddRange(GetFoldings(text, lowerCaseText, foldableKeyword));
            }


            return foldings.OrderBy((f) => f.StartOffset);
        }
        
        public void UpdateFoldings(FoldingManager manager, TextDocument document)
        {
            int firstErrorOffset;
            IEnumerable<NewFolding> foldings = CreateNewFoldings(document, out firstErrorOffset);
            manager.UpdateFoldings(foldings, firstErrorOffset);
        }
        

        public IEnumerable<NewFolding> GetFoldings(string text, string lowerCaseText, string keyword)
        {
            var foldings = new List<NewFolding>();

            string sPattern = $@"(\t|\f|\b| )+(((public|private).+|)({string.Join("|", foldablewords)}) +(.+))";
            string sPattern2 = $@"(\t|\f|\b| )+(((({string.Join("|", foldablewords)})))(.+))";
            string ePattern = $@"end(\t|\f|\b| )+({string.Join("|", foldablewords)})";
            string commPattern = $@"^(\t|\f|\b| |)+'(.+|)$";
            
            var stacks = new Stack<Tuple<Match, int, LineType>>();

            int index = 0, counter = 0;

            string[] lines = text.ToLower().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            
            foreach (string line in lines)
            {
                counter++;

                if (Regex.IsMatch(line, sPattern))
                {
                    Match m = Regex.Match(line, sPattern);

                    stacks.Push(new Tuple<Match,int, LineType>(m, index, LineType.Method));
                }
                else if (Regex.IsMatch(line, sPattern2))
                {
                    Match m = Regex.Match(line, sPattern2);

                    stacks.Push(new Tuple<Match, int, LineType>(m, index, LineType.Method));
                }
                else if (Regex.IsMatch(line, ePattern))
                {
                    Match Match = Regex.Match(line, ePattern);

                    if (stacks.Count == 0) goto Last;

                    var stack = stacks.Peek();
                    Match sMatch = stack.Item1;
                    
                    string eType = Match.Groups[2].Value;
                    string sType = sMatch.Groups[5].Value;

                    if (eType == sType) stacks.Pop();
                    else goto Last;

                    

                    var newFolding = new NewFolding(stack.Item2 + sMatch.Groups[2].Index, index + Match.Index + Match.Length);
                    newFolding.Name = text.Substring(stack.Item2 + sMatch.Groups[2].Index, sMatch.Groups[2].Length) + " ...";
                    foldings.Add(newFolding);
                }
                else if (Regex.IsMatch(line, commPattern))
                {
                    Match m = Regex.Match(line, commPattern);
                    
                    string value = m.Groups[2].Value;
                    

                    if (stacks.Count != 0)
                    {
                        var stack = stacks.Peek();
                        if (stack.Item3 == LineType.Method) goto Last;
                    }
                    

                    stacks.Push(new Tuple<Match, int, LineType>(m, index, LineType.Remark));
                }
                

                if (counter == lines.Count() || !Regex.IsMatch(line, commPattern))
                {
                    if (stacks.Count != 0)
                    {
                        int sIndex = 0, eIndex = 0;
                        bool flag = false;
                        string name = "";

                        int innerCounter = 0;
                        do
                        {
                            if (stacks.Count == 0)
                            {
                                if (innerCounter == 1) flag = false;

                                break;
                            }

                            var stack = stacks.Peek();
                            if (stack.Item3 != LineType.Remark) break;

                            stack = stacks.Pop();

                            sIndex = stack.Item2;
                            // 처음에만 endIndex 설정 나중에는 건들지 않기
                            if (!flag)
                                eIndex = stack.Item2 + stack.Item1.Value.Length;

                            name = stack.Item1.Groups[2].Value + " ...";

                            flag = true;

                            innerCounter++;

                        } while (true);

                        if (flag)
                        {
                            var newFolding = new NewFolding(sIndex, eIndex);
                            newFolding.Name = name;
                            foldings.Add(newFolding);
                        }
                    }
                }

                Last:
                index += line.Length + Environment.NewLine.Length;
            }
            return foldings;

        }
        public IEnumerable<NewFolding> GetFoldings_old(string text, string lowerCaseText, string keyword)
        {
            var foldings = new List<NewFolding>();

            string eKeyword = "end " + keyword;

            keyword += " ";

            var sOffsets = new Stack<int>();

            for (int i = 0; i <= text.Length - eKeyword.Length; i++)
            {
                if (lowerCaseText.Substring(i, keyword.Length) == keyword)
                {
                    int k = i;
                    if (k != 0)
                    {
                        while (text[k - 1].ToString() == "\r" || text[k - 1].ToString() == "\n")
                        {
                            k -= 1;
                            if (k <= 0) break;
                        }
                    }
                    sOffsets.Push(k);
                }
                else if (lowerCaseText.Substring(i, eKeyword.Length) == eKeyword)
                {
                    int sOffset = sOffsets.Pop();
                    var newFolding = new NewFolding(sOffset, i + eKeyword.Length);
                    int p = text.IndexOf("\r", sOffset);
                    if (p == -1) p = text.IndexOf("\n", sOffset);
                    if (p == -1) p = text.Length - 1;
                    newFolding.Name = text.Substring(sOffset, p - sOffset) + " ...";
                    foldings.Add(newFolding);
                }
                
            }


            return foldings;
        }

    }
}
