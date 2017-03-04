using Microsoft.Office.Core;
using ppt = Microsoft.Office.Interop.PowerPoint;
using PowerVBA.Core.Wrap.WrapClass;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using static PowerVBA.Core.Extension.boolEx;
using static PowerVBA.Global.Globals;
using System.Threading;
using VBA = Microsoft.Vbe.Interop;
using System.Text.RegularExpressions;

namespace PowerVBA.Core.Connector
{
    public class PPTConnector : IDisposable
    {

        #region [  전역 변수  ]
        public PresentationWrapping Presentation { get; internal set; }
        public ApplicationWrapping PowerPointApp { get; internal set; }

        public VBProjectWrapping VBProject { get; internal set; }

        Thread thr;

        #endregion

        #region [  생 성 자  ]

        public PPTConnector(string FileLocation,bool NewFile = false, bool OpenWithWindow = true) : this()
        {
            Presentation = new PresentationWrapping(PowerPointApp.Presentations.Open(FileLocation, MsoTriState.msoFalse, NewFile.BoolToState(), OpenWithWindow.BoolToState()));
            VBProject = new VBProjectWrapping(Presentation.VBProject);
            thr.Start();
        }
        public PPTConnector(bool OpenWithWindow = true) : this()
        {
            Presentation = new PresentationWrapping(PowerPointApp.Presentations.Add(OpenWithWindow.BoolToState()));
            VBProject = new VBProjectWrapping(Presentation.VBProject);
            Presentation.Slides.AddSlide(1, Presentation.SlideMaster.CustomLayouts[1]);

            thr.Start();
        }
        public PPTConnector()
        {
            PowerPointApp = new ApplicationWrapping(new ppt.Application());
            thr = new Thread(() => {
                do
                {
                    Thread.Sleep(100);
                    try { string name = Presentation.Presentation.Name; }
                    catch (Exception ex)
                    {
                        // 확인된 버전 : PPT 2013
                        // 프로그램이 사용중인 상태입니다. (인식되는 곳 : 종료 요청 등)
                        if (ex.HResult == -2147417846) continue;
                        PPTClosed();
                    }
                } while (true);
            });
        }

        #endregion


        public bool IsContainsName(string name)
        {
            foreach (VBA.VBComponent comp in VBProject.VBComponents)
            {
                if (name == comp.Name) return true;
            }

            return false;
        }

        public bool AddModule(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Var.name)) return false;
            if (IsContainsName(name)) return false;
            VBComponentWrapping newStandardModule = new VBComponentWrapping(VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule));

            newStandardModule.Name = name;

            return true;
        }

        public bool AddClass(string name)
        {
            if (!Regex.IsMatch(name, RegexPattern.Var.name)) return false;
            if (IsContainsName(name)) return false;
            VBComponentWrapping newStandardClass = new VBComponentWrapping(VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_ClassModule));

            newStandardClass.Name = name;

            foreach(VBA.Property d in newStandardClass.Properties)
            {
                Console.WriteLine(d.Name + " :: " + d.Value);
                if (d.Name == "Instancing")
                {
                    Console.WriteLine("Instancing Set To 2");
                    d.Value = 2;
                }
                
            }
            //newStandardClass

            return true;
        }



        public event BlankEventHandler PPTClosed;

        public void OnPPTClosed()
        {
            if (PPTClosed != null) PPTClosed();
        }


        public void Dispose()
        {
            Presentation.Close();
            if (PowerPointApp.Presentations.Count == 0) PowerPointApp.Quit();
        }
    }
}
