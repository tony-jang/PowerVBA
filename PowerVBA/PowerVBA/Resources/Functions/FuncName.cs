using PowerVBA.Core.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Resources.Functions
{
    public struct FuncName
    {
        public FuncName(string Folder, string FileName)
        {
            this.Folder = Folder;
            this.FileName = FileName;
        }

        public string Folder;
        public string FileName;

        public override string ToString()
        {
            return $"{Folder}.{FileName}";
        }

        public static implicit operator string(FuncName v)
        {
            if (string.IsNullOrEmpty(v.FileName)) throw new ArgumentNullException("파일 이름이 없습니다.");
            if (string.IsNullOrEmpty(v.Folder)) throw new ArgumentNullException("폴더 이름이 없습니다.");

            return $"{v.Folder}.{v.FileName}";
        }
        public static implicit operator FuncName(string s)
        {
            string[] sArr = s.Split(new char[] { '.' });
            if (sArr.Length != 2) throw new IndexOutOfRangeException("'.' 으로 Split 했을때 두개로 나누어지는 이름만 FuncName으로 변환 가능합니다.");
            
            return new FuncName(sArr[0], sArr[1]);
        }
    }
}
