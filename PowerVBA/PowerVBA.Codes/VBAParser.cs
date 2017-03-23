using ICSharpCode.AvalonEdit.Document;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public class VBAParser
    {

        public CodeNode Parse(string code)
        {
            return Parse(new List<string> { code });
        }

        public CodeNode Parse(List<string> codes)
        {
            CodeNode ReturnNode = new CodeNode();
            foreach (string code in codes)
            {
                CodeNode LineNode;
                RangeInt line = 0;
                int counter = 0;
                bool IsMultiLine = false;
                string data = "";

                foreach(string codesp in code.Split(new string[] { Environment.NewLine}, StringSplitOptions.None))
                {
                    counter++;
                    string codeline = codesp.Trim();
                    // 처리 - 다중 라인 처리
                    if (codeline.EndsWith("_"))
                    {
                        if (IsMultiLine)
                        {
                            line.EndInt = counter;
                            data += Environment.NewLine + codesp;
                            IsMultiLine = true;
                        }
                        else
                        {
                            line = counter;
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
                        }
                        else
                        // 처리 - 일반 적인 처리
                        {

                        }
                    }
                            

                }
            }


            return ReturnNode;
        }
    }
    

}
