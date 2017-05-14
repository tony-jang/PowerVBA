using ICSharpCode.AvalonEdit.Document;
using Microsoft.Vbe.Interop;
using PowerVBA.Codes.Attributes;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes
{
    public class VBAParser
    {
        CodeInfo CodeInfo;
        LineInfo LineInfo;

        public static List<string> VBANamespaces = new List<string>();
        public static void AddNameSpace(Assembly asm)
        {
            Type[] itm = asm.GetTypes()
                .Where(t => t.IsPublic)
                .Where(t => !t.Name.StartsWith("_"))
                .ToArray();

            foreach (Type type in itm)
            {
                if (VBANamespaces.Contains(type.Namespace)) VBANamespaces.Add(type.Namespace);
                if (type.IsInterface)
                {
                    //MessageBox.Show(type.ToString() + " :: Interface");
                }
                else if (type.IsClass && type.IsAbstract)
                {
                    //MessageBox.Show(type.ToString() + " :: Abstract Class");
                }
            }
        }

        static VBAParser()
        {
            AddNameSpace(Assembly.Load(Properties.Resources.LibPowerPoint));
            AddNameSpace(Assembly.Load(Properties.Resources.Interop_VBA));
        }
        public VBAParser(CodeInfo codeInfo)
        {
            this.CodeInfo = codeInfo;
            
            CodeInfo.Reset();
            
        }
        public void Parse((string, string) code)
        {
            Parse(new List<(string, string)> { code });
        }

        public CodeFileInfo GetFileInfo(string FileName)
        {
            // TODO : 구현
            return null;
        }

        public void Parse(List<(string, string)> codes)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            CodeInfo.ErrorList.Clear();
            CodeInfo.Lines.Clear();
            
            foreach ((string, string) code in codes)
            {
                LineInfo = new LineInfo();
                LineInfo.FileName = code.Item2;

                RangeInt CodeLine = 0;
                
                int LineCount = 0;
                bool IsMultiLine = false;
                string data = string.Empty;

                

                char[] cArr = code.Item1.ToCharArray();
                int Length = cArr.Length;
                
                string spCode = "";

                VBASeeker seeker = new VBASeeker(CodeInfo);

                
                for (int i=0;i< cArr.Length; i++)
                {
                    bool ReadLine = (cArr[i] == '\r' && cArr[i + 1] == '\n');
                    if (Length - 1 == i) {
                        ReadLine = true;
                        spCode += cArr[i];
                    }

                    if (ReadLine)
                    {
                        LineCount++;

                        string codeline = spCode.Trim();
                        // 처리 - 다중 라인 처리
                        if (codeline.EndsWith("_"))
                        {
                            if (IsMultiLine)
                            {
                                // 다중 줄 코드라면 끝나는 부분을 해당 라인으로 잡음
                                CodeLine.EndInt = LineCount;
                                data += Environment.NewLine + spCode;
                                    
                                IsMultiLine = true;
                            }
                            else
                            {
                                // 다중 줄의 첫 시작이라면 시작하는 부분을 해당 라인으로 잡음
                                CodeLine = LineCount;
                                data = spCode;
                                IsMultiLine = true;
                            }
                            continue;
                        }
                        else
                        {
                            NotHandledLine nHandledLine = NotHandledLine.Empty;
                            // 처리 - 이전 줄이 다중 라인이였을경우
                            if (IsMultiLine)
                            {
                                IsMultiLine = false;
                                nHandledLine = seeker.GetLine(code.Item2, spCode, (CodeLine, i), ref LineInfo);
                            }
                            // 처리 - 일반 적인 처리
                            else
                            {
                                CodeLine.StartInt = LineCount;
                                nHandledLine = seeker.GetLine(code.Item2, spCode, (CodeLine, i), ref LineInfo);
                            }

                            while (nHandledLine != NotHandledLine.Empty)
                            {
                                nHandledLine = seeker.GetLine(nHandledLine.FileName, nHandledLine.CodeLine, nHandledLine.Lines, ref LineInfo, true);
                            }
                        }
                        i++;
                        spCode = string.Empty;
                        continue;
                    }

                    spCode += cArr[i];
                }

                if (LineInfo.IsInFunction) AddError(ErrorCode.VB0210);
                else if (LineInfo.IsInSub) AddError(ErrorCode.VB0211);
                else if (LineInfo.IsInProperty) AddError(ErrorCode.VB0212);
                else if (LineInfo.IsInSelectCase) AddError(ErrorCode.VB0213);
                else if (LineInfo.IsInIf) AddError(ErrorCode.VB0214);
                else if (LineInfo.IsInDo) AddError(ErrorCode.VB0215);
                else if (LineInfo.IsInDoWhile) AddError(ErrorCode.VB0216);
                else if (LineInfo.IsInDoUntil) AddError(ErrorCode.VB0217);
                else if (LineInfo.IsInEnum) AddError(ErrorCode.VB0218);
                else if (LineInfo.IsInFor) AddError(ErrorCode.VB0219);
                else if (LineInfo.IsInForEach) AddError(ErrorCode.VB0210);

                void AddError(ErrorCode Code, string[] parameters = null)
                {
                    CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, code.Item2, new DomRegion(LineCount, 0)));
                }
                CodeInfo.Lines.Add(LineInfo);
            }

            GC.Collect();

            //Errors.Where((i) => i.ErrorCode.ContainAttribute(typeof(NotSupportedAttribute))).ToList().ForEach((i) => Errors.Remove(i));
            
        }


    }
    

}
