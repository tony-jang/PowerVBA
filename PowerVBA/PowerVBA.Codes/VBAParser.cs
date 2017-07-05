using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

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
        public void Parse((string, string) fileName)
        {
            Parse(new List<(string,string)>() { fileName });
        }

        public void Parse(List<(string, string)> codeFiles)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            var fileNames = codeFiles.Select(j => j.Item1);

            CodeInfo.ErrorList.Where(i => fileNames.Contains(i.FileName))
                .ToList()
                .ForEach(j => CodeInfo.ErrorList.Remove(j));

            CodeInfo.Lines.Where(i => fileNames.Contains(i.FileName))
                .ToList()
                .ForEach(j => CodeInfo.Lines.Remove(j));

            CodeInfo.CodeFiles.Where(i => fileNames.Contains(i.FileName))
                .ToList()
                .ForEach(j => 
                {
                    j.Variables.Clear();
                    j.Functions.Clear();
                    j.Subs.Clear();
                    j.Enums.Clear();
                });

            foreach ((string,string) codeFile in codeFiles)
            {
                LineInfo = new LineInfo();
                LineInfo.FileName = codeFile.Item1;
                
                RangeInt CodeLine = 0;
                
                int LineCount = 0;
                bool IsMultiLine = false;
                string data = string.Empty;

                char[] cArr = codeFile.Item2.ToCharArray();
                int Length = cArr.Length;
                
                string spCode = string.Empty;

                VBASeeker seeker = new VBASeeker(CodeInfo);

                
                for (int i = 0; i < cArr.Length; i++)
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
                                nHandledLine = seeker.GetLine(codeFile.Item1, spCode, CodeLine, ref LineInfo);
                            }
                            // 처리 - 일반 적인 처리
                            else
                            {
                                CodeLine.StartInt = LineCount;
                                nHandledLine = seeker.GetLine(codeFile.Item1, spCode, CodeLine, ref LineInfo);
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

                if (LineInfo.IsInFunction)
                    AddError(ErrorCode.VB0210, new string[] { "Function", "End Function" });
                else if (LineInfo.IsInSub)
                    AddError(ErrorCode.VB0210, new string[] { "Sub", "End Sub" });
                else if (LineInfo.IsInProperty)
                    AddError(ErrorCode.VB0210, new string[] { "Property", "End Property" });
                else if (LineInfo.IsInType)
                    AddError(ErrorCode.VB0210, new string[] { "Type", "End Type" });
                else if (LineInfo.IsInSelectCase)
                    AddError(ErrorCode.VB0210, new string[] { "Select Case", "End Select" });
                else if (LineInfo.IsInIf)
                    AddError(ErrorCode.VB0210, new string[] { "If", "End If" });
                else if (LineInfo.IsInDo)
                    AddError(ErrorCode.VB0210, new string[] { "Do", "Loop" });
                else if (LineInfo.IsInDoWhile)
                    AddError(ErrorCode.VB0210, new string[] { "Do While", "Loop" });
                else if (LineInfo.IsInDoUntil)
                    AddError(ErrorCode.VB0210, new string[] { "Do Until", "Loop" });
                else if (LineInfo.IsInEnum)
                    AddError(ErrorCode.VB0210, new string[] { "Enum", "End Enum" });
                else if (LineInfo.IsInFor)
                    AddError(ErrorCode.VB0210, new string[] { "For", "Next" });
                else if (LineInfo.IsInForEach)
                    AddError(ErrorCode.VB0210, new string[] { "For Each", "Next" });

                void AddError(ErrorCode Code, string[] parameters = null)
                {
                    CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, codeFile.Item1, LineCount));
                }
                CodeInfo.Lines.Add(LineInfo);
            }

            GC.Collect();
        }
    }
}
