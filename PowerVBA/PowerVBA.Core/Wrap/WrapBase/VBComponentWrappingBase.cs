using PowerVBA.Codes;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Interface;

namespace PowerVBA.Core.Wrap.WrapBase
{
    public abstract class VBComponentWrappingBase : IWrappingClass
    {
        public abstract PPTVersion ClassVersion { get; }
        public abstract string CompName { get; }
        public abstract string Code { get; set; }
        public CodeFile File { get; set; }

        public event BlankDelegate NameChanged;

        /// <summary>
        /// 0 : 알 수 없음 |
        /// 1 : 클래스 |
        /// 2 : 모듈 |
        /// 3 : 폼 |
        /// 4 : 슬라이드 클래스
        /// </summary>
        /// <returns></returns>
        public abstract int GetComponentType();
        public abstract string GetExtension { get; }

        public abstract void SetCode(string code);

        public abstract void CodeClear();

        public VBComponentWrappingBase(string Name)
        {
            this.NameChanged += VBComponentWrappingBase_NameChanged;
            File = new CodeFile(Name);
        }

        public void OnNameChanged()
        {
            NameChanged?.Invoke();
        }

        private void VBComponentWrappingBase_NameChanged()
        {
            File.FileName = CompName;
        }
    }
}