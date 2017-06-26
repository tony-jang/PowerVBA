using PowerVBA.Codes.Extension;
using PowerVBA.Codes.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes
{
    public class CodeFile
    {
        private CodeFile()
        {
            Functions = new List<Function>();
            Subs = new List<Sub>();
            Properties = new List<Property>();
            Variables = new VariableManager();
            Enums = new List<EnumItem>();
        }

        public CodeFile(string fileName) : this()
        {
            this.FileName = fileName;
        }

        /// <summary>
        /// 파일 이름입니다.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 함수 목록입니다.
        /// </summary>
        public List<Function> Functions { get; internal set; }
        /// <summary>
        /// 서브루틴 목록입니다.
        /// </summary>
        public List<Sub> Subs { get; internal set; }
        /// <summary>
        /// 프로퍼티 목록입니다.
        /// </summary>
        public List<Property> Properties { get; internal set; }
        /// <summary>
        /// 변수 목록입니다.
        /// </summary>
        public VariableManager Variables { get; internal set; }

        public List<EnumItem> Enums { get; internal set; }

        public bool MemberContains(string memberName)
        {
            List<IMember> members = Functions
                .Select(i => (IMember)i)
                .Concat(Subs.Select(i => (IMember)i))
                .Concat(Variables.Select(i => (IMember)i))
                .Concat(Enums.Select(i => (IMember)i))
                .ToList();

            return members
                .Select(i => i.Name)
                .Contains(memberName);
        }

        /// <summary>
        /// 코드 변수를 추가합니다.
        /// </summary>
        public bool AddVaraible(string name, string fileName, LinePoint point)
        {
            if (Variables.Where(i => i.Name == name)
                         .Count() != 0)
                return false;

            Variables.Add(new Variable(name, fileName, point.Line, point.Offset));
            return true;
        }

        /// <summary>
        /// 코드 변수를 제거합니다.
        /// </summary>
        public bool RemoveVariable(string name)
        {
            var itm = Variables.Where(i => i.Name == name).FirstOrDefault();
            if (itm == null) return false;

            Variables.Remove(itm);

            return true;
        }

        /// <summary>
        /// 변수 사용 위치를 추가합니다.
        /// </summary>
        public bool AddVariableUsage(string name, int line, int offset)
        {
            var itm = Variables.Where(i => i.Name == name).FirstOrDefault();
            if (itm == null) return false;

            itm.Usages.Add((line, offset, name.Length));

            return true;
        }

    }
}
