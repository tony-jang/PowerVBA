using Microsoft.Win32;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerVBA
{
    // MainWindow Partial File :: Initalize Window Part
    // MainWindow 부분 파일 :: 초기화 윈도우 부분
    
    partial class MainWindow
    {
        private void GridOpenAnotherPpt_MouseLeave(object sender, MouseEventArgs e)
        {
            TextBlock tb = null;
            if (((Control)sender).Name.ToLower().Contains("openanotherppt")) tb = tbBOpenAnotherPpt;
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation")) tb = tbConnectPresentation;

            if (tb != null)
            {
                tb.FontStyle = FontStyles.Normal;
                while (tb.TextDecorations.Count != 0)
                {
                    tb.TextDecorations.RemoveAt(0);
                }
            }
        }

        private void OpenButtons_MouseMove(object sender, MouseEventArgs e)
        {
            TextBlock tb = null;
            if (((Control)sender).Name.ToLower().Contains("openanotherppt")) tb = tbBOpenAnotherPpt;
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation")) tb = tbConnectPresentation;

            tb?.TextDecorations.Add(TextDecorations.Underline);
        }

        private void OpenButtons_Click(object sender, RoutedEventArgs e)
        {
            if (((Control)sender).Name.ToLower().Contains("openanotherppt"))
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    Filter = "프레젠테이션|*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm"
                };
                if (ofd.ShowDialog().Value)
                {
                    tbProcessInfoTB.Text = "프레젠테이션을 열고 있습니다.";

                    dbConnector.RecentFile_Add(ofd.FileName);

                    InitalizeConnector(ofd.FileName);
                }
            }
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation"))
            {
                ConnectWindows connWindow = new ConnectWindows();

                var Handled = connWindow.ShowDialog(out PresentationWrappingBase ppt);

                if (Handled)
                {
                    try
                    {
                        InitalizeConnector(ppt);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        private void BtnNewPPT_Click(object sender, RoutedEventArgs e)
        {

            PPTVersion ver = VersionSelector.GetPPTVersion();

            if (ver != PPTVersion.PPT2013 && ver != PPTVersion.PPT2010 && ver != PPTVersion.PPT2016)
            {
                MessageBox.Show($"죄송합니다. {ver.ToString()}는 지원하지 않는 버전입니다.");
                return;
            }
            tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

            InitalizeConnector();
        }

        private void SelectionChangedDetect()
        {
            projAnalyzer.CurrentShapeName = connector.SelectionShapeName;
        }

        private void ShapeChangedDetect()
        {
            var shapeCount = connector.ShapeCount;
            connector.AutoShapeUpdate = !(shapeCount < 1000);
            projAnalyzer.ShapeCount = shapeCount;
            projAnalyzer.CurrentShapeCount = connector.Shapes(connector.Slide).Count();
        }

        private void SlideChangedDetect()
        {
            projAnalyzer.SlideCount = connector.SlideCount;
        }
    }
}
