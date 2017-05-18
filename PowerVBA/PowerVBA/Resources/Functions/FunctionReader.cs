using PowerVBA.Core.Extension;
using PowerVBA.Core.Global;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerVBA.Resources.Functions
{
    /// <summary>
    /// 미리 정의된 함수들을 관리하는 클래스입니다.
    /// </summary>
    public static class FunctionReader
    {
        static Dictionary<string, Function> funcDic = new Dictionary<string, Function>();
        public static List<Function> GetFunctions()
        {
            List<Function> funcList = new List<Function>();
            foreach ((string, string) s in ResourceManager.GetCodeDatas("Functions"))
            {
                string folderPath = "PowerVBA.Resources.Functions.";
                string Path = s.Item1.Substring(folderPath.Length);
                Path = Path.Substring(0, Path.Length - 4);
                funcList.Add(GetFunction(Path));
            }
            return funcList;
        }
        
        public static Function GetFunction(string FileName)
        {
            string code = ResourceManager.GetTextResource($"Functions/{FileName}.txt");
            string[] codeArr = code.SplitByNewLine();

            string attrPattern = "^'(.+?):(.+)$";

            Regex r = new Regex(attrPattern);
            int counter = 0;

            string description = "",
                   dependency = "",
                   returnType = "";

            do
            {
                Match m = r.Match(codeArr[counter]);
                if (m.Success)
                {
                    string attrName = m.Groups[1].Value;
                    string value = m.Groups[2].Value;

                    var attr = Enum.GetValues(typeof(FuncAttributes))
                        .Cast<FuncAttributes>()
                        .Where(i => i.GetDescription() == attrName);
                    if (attr.Count() == 0)
                    {
                        ErrorBox.Show($"{FileName}에서 발견된 {attrName} 특성은 미리 지정된 함수 특성에 포함되지 않습니다. 현재 파일의 읽기 작업을 건너뜁니다.", ErrorType.Resource);
                        continue;
                    }

                    FuncAttributes attribute = attr.First();

                    switch (attribute)
                    {
                        case FuncAttributes.ReturnType:
                            returnType = value;
                            break;
                        case FuncAttributes.Description:
                            description = value;
                            break;
                        case FuncAttributes.Dependency:
                            dependency = value;
                            break;
                    }
                }
                else
                {
                    break;
                }
                counter++;
            } while (true);
            

            if (string.IsNullOrEmpty(code)) return null;

            FuncName name = FileName;

            var f = new Function(FileName, 
                                 string.Join(Environment.NewLine, codeArr.SubArray(counter)),
                                 !string.IsNullOrEmpty(returnType), 
                                 returnType, 
                                 dependency,description);
            if (!funcDic.ContainsKey(FileName))
            {
                funcDic[FileName] = f;
            }
            return funcDic[FileName];
        }
    }
}
