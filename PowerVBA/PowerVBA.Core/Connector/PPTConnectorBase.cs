using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap.WrapBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Connector
{
    public abstract class PPTConnectorBase : IPPTConnector
    {
        public PresentationWrappingBase _Presentation { get; set; }
        public ApplicationWrappingBase _PPTApp { get; set; }
        public VBProjectWrappingBase _VBProject { get; set; }

        public abstract PPTVersion Version { get; }

        public event BlankDelegate VBAComponentChange;
        public event BlankDelegate PresentationClosed;
        public event BlankDelegate ShapeChanged;
        public event BlankDelegate SlideChanged;
        public event BlankDelegate SectionChanged;
        
        public void OnVBAComponentChange() { VBAComponentChange?.Invoke(); }
        public void OnPPTClosed() { PresentationClosed?.Invoke(); }
        public void OnShapeChanged() { ShapeChanged?.Invoke(); }
        public void OnSlideChanged() { SlideChanged?.Invoke(); }
        public void OnSectionChanged() { SectionChanged?.Invoke(); }
        
        



        public abstract List<ShapeWrappingBase> Shapes();



        /// <summary>
        /// 슬라이드를 추가합니다.
        /// </summary>
        /// <returns></returns>
        public abstract bool AddSlide();

        /// <summary>
        /// 슬라이드를 추가합니다.
        /// </summary>
        /// <param name="SlideNumber">추가할 슬라이드 번호를 입력합니다.</param>
        /// <returns></returns>
        public abstract bool AddSlide(int SlideNumber);

        /// <summary>
        /// 슬라이드를 삭제합니다.
        /// </summary>
        /// <returns></returns>
        public abstract bool DeleteSlide();

        /// <summary>
        /// 슬라이드를 삭제합니다.
        /// </summary>
        /// <param name="SlideNumber">삭제할 슬라이드 번호를 입력합니다.</param>
        /// <returns></returns>
        public abstract bool DeleteSlide(int SlideNumber);


        /// <summary>
        /// 모듈을 추가합니다.
        /// </summary>
        /// <param name="name">추가할 모듈 이름입니다.</param>
        /// <returns></returns>
        public abstract bool AddModule(string name);
        /// <summary>
        /// 모듈을 제거합니다.
        /// </summary>
        /// <param name="name">제거할 모듈 이름입니다.</param>
        /// <returns></returns>
        public abstract bool DeleteModule(string name);

        /// <summary>
        /// 클래스를 추가합니다.
        /// </summary>
        /// <param name="name">추가할 클래스 이름입니다.</param>
        /// <returns></returns>
        public abstract bool AddClass(string name);
        /// <summary>
        /// 클래스를 제거합니다.
        /// </summary>
        /// <param name="name">제거할 클래스 이름입니다.</param>
        /// <returns></returns>
        public abstract bool DeleteClass(string name);

        /// <summary>
        /// 폼을 추가합니다.
        /// </summary>
        /// <param name="name">추가할 폼 이름입니다.</param>
        /// <returns></returns>
        public abstract bool AddForm(string name);
        /// <summary>
        /// 폼을 제거합니다.
        /// </summary>
        /// <param name="name">제거할 폼 이름입니다.</param>
        /// <returns></returns>
        public abstract bool DeleteForm(string name);

        public abstract void Dispose();
    }
}
