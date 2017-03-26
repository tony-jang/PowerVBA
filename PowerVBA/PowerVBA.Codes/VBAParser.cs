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
        public List<Error> Errors;
        public VBAParser(List<Error> errors)
        {
            Errors = errors;
            Errors.Clear();
        }
        public List<CodeData> Parse(string code)
        {
            return Parse(new List<string> { code });
        }


        List<CodeData> codeDatas = new List<CodeData>();

        public List<CodeData> Parse(List<string> codes)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            int counter = 0;

            LineInfo lineInfo = new LineInfo();

            foreach (string code in codes)
            {
                RangeInt line = 0;
                
                int LineCount = 0;
                bool IsMultiLine = false;
                string data = "";

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
                            line.EndInt = LineCount;
                            data += Environment.NewLine + codesp;
                            IsMultiLine = true;
                        }
                        else
                        {
                            line = LineCount;
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
                            VBASeeker seeker = new VBASeeker(data + Environment.NewLine + codesp, line, Errors, lineInfo);
                            codeDatas = codeDatas.Concat(seeker.GetLine()).ToList();
                            //MessageBox.Show(string.Join("\r\n", seeker.GetLine().ChildNode.Select((i) => i.ToString())));
                        }
                        else
                        // 처리 - 일반 적인 처리
                        {
                            line.StartInt = LineCount;
                            VBASeeker seeker = new VBASeeker(codesp, line, Errors, lineInfo);
                            codeDatas = codeDatas.Concat(seeker.GetLine()).ToList();
                            //MessageBox.Show(string.Join("\r\n", seeker.GetLine().ChildNode.Select((i) => i.ToString())));
                        }
                    }

                    
                }
                
            }

            // DEBUG: 지원하지 않는 문법은 제외
            //Errors.Where((i) => i.ErrorCode.ContainAttribute(typeof(NotSupportedAttribute))).ToList().ForEach((i) => Errors.Remove(i));
                
            
            return codeDatas;
        }
    }
    

}
