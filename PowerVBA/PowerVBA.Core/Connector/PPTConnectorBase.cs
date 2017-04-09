using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap.WrapBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PowerVBA.Global.Globals;

namespace PowerVBA.Core.Connector
{
    public abstract class PPTConnectorBase : IPPTConnector
    {
        public PresentationWrappingBase _Presentation { get; set; }
        public ApplicationWrappingBase _PPTApp { get; set; }
        public VBProjectWrappingBase _VBProject { get; set; }

        public abstract PPTVersion Version { get; }
        public abstract string Name { get; }

        public event BlankDelegate VBAComponentChange;
        public event BlankDelegate PresentationClosed;
        public event BlankDelegate ShapeChanged;
        public event BlankDelegate SlideChanged;
        public event BlankDelegate SectionChanged;
        
        public void OnVBAComponentChange()
        {
            try { MainDispatcher.Invoke(new Action(() => { VBAComponentChange?.Invoke(); })); } catch (Exception) { }    
        }
        public void OnPPTClosed()
        {
            try { MainDispatcher.Invoke(new Action(() => { PresentationClosed?.Invoke(); })); } catch (Exception) {}
        }
        public void OnShapeChanged()
        {
            try { MainDispatcher.Invoke(new Action(() => { ShapeChanged?.Invoke(); })); } catch (Exception) { }
        }
        
        public void OnSlideChanged()
        {
            try { MainDispatcher.Invoke(new Action(() => { SlideChanged?.Invoke(); })); } catch (Exception) { }    
        }
        
        public void OnSectionChanged()
        {
            try { MainDispatcher.Invoke(new Action(() => { SectionChanged?.Invoke(); })); } catch (Exception) { }
        }





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
        public abstract bool AddModule(string name, out VBComponentWrappingBase comp);
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
        public abstract bool AddClass(string name, out VBComponentWrappingBase comp);
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
        public abstract bool AddForm(string name, out VBComponentWrappingBase comp);
        /// <summary>
        /// 폼을 제거합니다.
        /// </summary>
        /// <param name="name">제거할 폼 이름입니다.</param>
        /// <returns></returns>
        public abstract bool DeleteForm(string name);

        public abstract void Dispose();
        /// <summary>
        /// 해당 이름에 해당하는 모듈이 존재하는지에 대한 여부를 가져옵니다.
        /// </summary>
        /// <param name="name">확인할 모듈 이름입니다.</param>
        /// <returns></returns>
        public abstract bool ContainsModule(string name);
        /// <summary>
        /// 해당 이름에 해당하는 클래스가 존재하는지에 대한 여부를 가져옵니다.
        /// </summary>
        /// <param name="name">확인할 클래스 이름입니다.</param>
        /// <returns></returns>
        public abstract bool ContainsClass(string name);

        /// <summary>
        /// 해당 이름에 해당하는 사용자 지정 폼이 존재하는지에 대한 여부를 가져옵니다.
        /// </summary>
        /// <param name="name">확인할 사용자 지정 폼의 이름입니다.</param>
        /// <returns></returns>
        public abstract bool ContainsForm(string name);
        
        public abstract VBComponentWrappingBase GetModule(string name);

        public abstract VBComponentWrappingBase GetClass(string name);

        public abstract VBComponentWrappingBase GetFrm(string name);

        public abstract bool Save();
        public abstract bool SaveAs(string path);
    }
}
