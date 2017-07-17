using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    /// <summary>
    /// 현재 프로젝트 코드에 대한 모든 정보를 담고 있습니다.
    /// </summary>
    public class CodeInfo
    {

        public CodeInfo()
        {
            ErrorList = new List<Error>();
            Lines = new List<LineInfo>();

            CodeFiles = new List<CodeFile>();
        }

        #region [  Find Members  ]

        public Variable FindVariable(string file, string name)
        {
            var codeFile = CodeFiles.Where(i => i.FileName == file).FirstOrDefault();

            if (codeFile != null)
            {
                var variable = codeFile.Variables.Where(i => i.Name == name).FirstOrDefault();

                return variable;
            }

            return null;
        }

        public Function FindFunction(string file, string name)
        {
            var codeFile = CodeFiles.Where(i => i.FileName == file).FirstOrDefault();

            if (codeFile != null)
            {
                var function = codeFile.Functions.Where(i => i.Name == name).FirstOrDefault();

                return function;
            }

            return null;
        }

        #endregion

        /// <summary>
        /// 현재 프로젝트 코드의 오류를 나타냅니다.
        /// </summary>
        public List<Error> ErrorList { get; set; }
        
        public List<LineInfo> Lines { get; set; }
        
        public List<CodeFile> CodeFiles { get; set; }

        /// <summary>
        /// 파일을 추가합니다. 이미 있는 경우는 덮어씁니다.
        /// 덮어쓰기를 허용하고 싶지 않다면 overrides를 false로 선택해주세요.
        /// </summary>
        /// <param name="fileName"></param>
        public void AddFile(string fileName, bool overrides = true)
        {
            var itm = CodeFiles.Where(i => i.FileName == fileName);
            if (itm.Count() != 0 && overrides)
            {
                itm.ToList()
                   .ForEach(i => CodeFiles.Remove(i));
            }
            CodeFiles.Add(new CodeFile(fileName));
        }

        public void AddFile(CodeFile codeFile, bool overrides = true)
        {
            if (codeFile == null)
                return;
            var itm = CodeFiles.Where(i => i.FileName == codeFile.FileName);
            if (itm.Count() != 0 && overrides)
            {
                itm.ToList()
                   .ForEach(i => CodeFiles.Remove(i));
            }
            CodeFiles.Add(codeFile);
        }

        public CodeFile GetFile(string fileName)
        {
            return CodeFiles.Where(i => i.FileName == fileName).FirstOrDefault();
        }

        public bool RemoveFile(string fileName)
        {
            var itm = CodeFiles.Where(i => i.FileName == fileName);

            if (itm.FirstOrDefault() != null)
            {
                CodeFiles.Remove(itm.FirstOrDefault());
                return true;
            }
            return false;
        }

        public int FunctionCount
        {
            get
            {
                int count = 0;
                foreach (var codeFile in CodeFiles)
                {
                    count += codeFile.Functions.Count;
                }
                return count;
            }
        }
        public int SubCount
        {
            get
            {
                int count = 0;
                foreach (var codeFile in CodeFiles)
                {
                    count += codeFile.Subs.Count;
                }
                return count;
            }
        }
        public int PropertyCount
        {
            get
            {
                int count = 0;
                foreach (var codeFile in CodeFiles)
                {
                    count += codeFile.Properties.Count;
                }
                return count;
            }
        }

        public void Reset()
        {
            ErrorList = new List<Error>();
        }

        public override string ToString()
        {
            List<string> returnD = new List<string>();
            
            return string.Join("\r\n", returnD.ToArray());
        }
    }
}
