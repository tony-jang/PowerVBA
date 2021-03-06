﻿using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Core.Error;
using PowerVBA.Core.Project;
using PowerVBA.RegexPattern;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static PowerVBA.Core.Extension.StringEx;
using System.Collections;
using System.Windows;
using PowerVBA.Core.AvalonEdit.CodeCompletion;
using System.Diagnostics;

namespace PowerVBA.Core.AvalonEdit.Parser
{
    public class CodeParser : IList<ILineInfo>
    {
        public CodeParser(TextEditor editor, List<CodeError> errors)
        {
            this.Editor = editor;
            this.Errors = errors;
        }

        public TextEditor Editor { get; }
        public List<CodeError> Errors { get; }

        
        public List<ILineInfo> Lines { get; }
        
        public long Seek()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            DocumentLine currLine = Editor.Document.GetLineByOffset(Editor.CaretOffset);

            string currCode = Editor.Text.Substring(currLine.Offset, Editor.CaretOffset - currLine.Offset);
            
            if (currCode.Length >= 2)
            {
                string PrevText = currCode.Substring(Editor.CaretOffset - currLine.Offset - 2 , 1);
                string ThisText = currCode.Substring(Editor.CaretOffset - currLine.Offset - 1, 1);

                if (PrevText == " " && ThisText == " ")
                    return sw.ElapsedMilliseconds;

            }

            MiniLexer lexer = new MiniLexer(currCode);

            lexer.Parse();

            if (lexer.IsInSingleComment || lexer.IsInString) return sw.ElapsedMilliseconds;


            //// 라인 분석 (문법적 선언 오류)
            foreach (DocumentLine line in Editor.Document.Lines)
            {
                string code = Editor.Text.Substring(line.Offset, line.Length);

                LineParser lparser = new LineParser(code, this, new AnchorSegment(Editor.Document, line.Offset, line.Length));
                var itm = lparser.Parse();
                //MessageBox.Show(itm.LineType.ToString());
            }

            return sw.ElapsedMilliseconds;
        }



        #region [  Interface Implements  ]

        public IList<ILineInfo> Lists = new List<ILineInfo>();

        public ILineInfo this[int index]
        {
            get { return Lists[index]; }
            set { Lists[index] = value; }
        }

        public int Count
        {
            get { return Lists.Count; }
        }

        public bool IsReadOnly
        {
            get { return Lists.IsReadOnly; }
        }

        public void Add(ILineInfo item)
        {
            Lists.Add(item);
        }

        public void Clear()
        {
            Lists.Clear();
        }

        public bool Contains(ILineInfo item)
        {
            return Lists.Contains(item);
        }

        public void CopyTo(ILineInfo[] array, int arrayIndex)
        {
            Lists.CopyTo(array, arrayIndex);
        }

        public IEnumerator<ILineInfo> GetEnumerator()
        {
            return Lists.GetEnumerator();
        }

        public int IndexOf(ILineInfo item)
        {
            return Lists.IndexOf(item);
        }

        public void Insert(int index, ILineInfo item)
        {
            Lists.Insert(index, item);
        }

        public bool Remove(ILineInfo item)
        {
            return Lists.Remove(item);
        }

        public void RemoveAt(int index)
        {
            Lists.RemoveAt(index);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return Lists.GetEnumerator();
        }

        #endregion


    }
}
