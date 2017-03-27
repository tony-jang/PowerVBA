using ICSharpCode.AvalonEdit.Document;
using Microsoft.Vbe.Interop;
using PowerVBA.Codes.Attributes;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes
{
    public class VBAParser
    {
        CodeInfo CodeInfo;
        public VBAParser(CodeInfo codeInfo)
        {
            this.CodeInfo = codeInfo;
        }
        public void Parse(string code)
        {
            Parse(new List<string> { code });
        }

        public void Parse(List<string> codes)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            int counter = 0;

            // 기존 코드의 네스트(포함 여부)를 확인하기 위한 변수
            LineInfo lineInfo = new LineInfo();

            foreach (string code in codes)
            {
                
                RangeInt CodeLine = 0;
                
                int LineCount = 0;
                int Offset = 0;
                bool IsMultiLine = false;
                string data = string.Empty;

                foreach(string codesp in code.Split(new string[] { Environment.NewLine}, StringSplitOptions.None))
                {
                    LineCount++;

                    counter++;
                    string codeline = codesp.Trim();
                    // 처리 - 다중 라인 처리
                    if (codeline.EndsWith("_"))
                    {
                        if (IsMultiLine)
                        {
                            // 다중 줄 코드라면 끝나는 부분을 해당 라인으로 잡음
                            CodeLine.EndInt = LineCount;
                            data += Environment.NewLine + codesp;
                            IsMultiLine = true;
                        }
                        else
                        {
                            // 다중 줄의 첫 시작이라면 시작하는 부분을 해당 라인으로 잡음
                            CodeLine = LineCount;
                            data = codesp;
                            IsMultiLine = true;
                        }
                        continue;
                    }
                    else
                    {
                        // 처리 - 이전 줄이 다중 라인이였을경우
                        if (IsMultiLine)
                        {
                            IsMultiLine = false;
                            new VBASeeker(CodeInfo).GetLine(data + Environment.NewLine + codesp, CodeLine);
                        }
                        // 처리 - 일반 적인 처리
                        else
                        {
                            CodeLine.StartInt = LineCount;
                            new VBASeeker(CodeInfo).GetLine(codesp, CodeLine);
                        }
                    }
                    
                    Offset += codesp.Length + Environment.NewLine.Length;
                }
                
            }
            
            //Errors.Where((i) => i.ErrorCode.ContainAttribute(typeof(NotSupportedAttribute))).ToList().ForEach((i) => Errors.Remove(i));
            
        }
    }
    

}
