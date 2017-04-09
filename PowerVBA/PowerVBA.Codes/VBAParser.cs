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



            //StringBuilder sb = new StringBuilder();
            //foreach(Type t in asm.GetTypes())
            //{
            //    sb.AppendLine(t.Namespace + "." + t.Name + $" ({(t.IsNotPublic ? "Private" : "Public" )})");
            //}
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
        public void Parse(string code)
        {
            Parse(new List<string> { code });
        }

        public void Parse(List<string> codes)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            CodeInfo.ErrorList.Clear();
            
            foreach (string code in codes)
            {
                RangeInt CodeLine = 0;
                
                int LineCount = 0;
                bool IsMultiLine = false;
                string data = string.Empty;

                char[] cArr = code.ToCharArray();
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
                            // 처리 - 이전 줄이 다중 라인이였을경우
                            if (IsMultiLine)
                            {
                                IsMultiLine = false;
                                seeker.GetLine("", spCode, (CodeLine, i));
                            }
                            // 처리 - 일반 적인 처리
                            else
                            {
                                CodeLine.StartInt = LineCount;
                                seeker.GetLine("", spCode, (CodeLine, i));
                            }
                        }

                        i++;
                        spCode = string.Empty;
                        continue;
                    }

                    spCode += cArr[i];
                }
            }
            
            //Errors.Where((i) => i.ErrorCode.ContainAttribute(typeof(NotSupportedAttribute))).ToList().ForEach((i) => Errors.Remove(i));
            
        }
    }
    

}
