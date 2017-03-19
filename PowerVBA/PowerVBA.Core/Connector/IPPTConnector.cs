using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap.WrapBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Connector
{
    public delegate void BlankDelegate();
    public interface IPPTConnector : IDisposable
    {
        
        PPTVersion Version { get; }



        /// <summary>
        /// VBA 프로젝트의 파일이 변경되었습니다.
        /// </summary>
        event BlankDelegate VBAComponentChange;
        /// <summary>
        /// PPT Prensentation이 종료되었습니다.
        /// </summary>
        event BlankDelegate PresentationClosed;
        /// <summary>
        /// Shape (도형) 이 변경되었습니다.
        /// </summary>
        event BlankDelegate ShapeChanged;
        /// <summary>
        /// Slide (슬라이드) 가 변경되었습니다.
        /// </summary>
        event BlankDelegate SlideChanged;
        /// <summary>
        /// Section (섹션) 이 변경되었습니다.
        /// </summary>
        event BlankDelegate SectionChanged;


        /// <summary>
        /// 모듈을 추가합니다.
        /// </summary>
        /// <param name="name">추가할 모듈 이름입니다.</param>
        /// <returns></returns>
        bool AddModule(string name, out VBComponentWrappingBase comp);
        /// <summary>
        /// 모듈을 제거합니다.
        /// </summary>
        /// <param name="name">제거할 모듈 이름입니다.</param>
        /// <returns></returns>
        bool DeleteModule(string name);
        
        /// <summary>
        /// 클래스를 추가합니다.
        /// </summary>
        /// <param name="name">추가할 클래스 이름입니다.</param>
        /// <returns></returns>
        bool AddClass(string name, out VBComponentWrappingBase comp);
        /// <summary>
        /// 클래스를 제거합니다.
        /// </summary>
        /// <param name="name">제거할 클래스 이름입니다.</param>
        /// <returns></returns>
        bool DeleteClass(string name);

        /// <summary>
        /// 폼을 추가합니다.
        /// </summary>
        /// <param name="name">추가할 폼 이름입니다.</param>
        /// <returns></returns>
        bool AddForm(string name, out VBComponentWrappingBase comp);
        /// <summary>
        /// 폼을 제거합니다.
        /// </summary>
        /// <param name="name">제거할 폼 이름입니다.</param>
        /// <returns></returns>
        bool DeleteForm(string name);


    }
}
