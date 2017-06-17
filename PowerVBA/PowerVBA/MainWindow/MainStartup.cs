using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Windows;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerVBA
{
    partial class MainWindow
    {
        // 도움말 열기
        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            var helperWindow = new HelperWindow();
            helperWindow.MoveHelpContext(((Button)sender).Tag.ToString());
            helperWindow.ShowDialog();
        }

        // 새로운 PPT 열기
        private void BtnNewPPT_Click(object sender, RoutedEventArgs e)
        {
            // 연결되어있는데 저장 여부에서 취소를 눌렀을시 처리 중지
            if (IsConnected && !CloseRequest()) return;

            InitalizeAll();
            PPTVersion ver = VersionSelector.GetPPTVersion();

            if (ver != PPTVersion.PPT2013 && ver != PPTVersion.PPT2010 && ver != PPTVersion.PPT2016)
            {
                MessageBox.Show($"죄송합니다. {ver.ToString()}는 지원하지 않는 버전입니다.");
                return;
            }
            tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

            InitalizeConnector();
        }

        private void GridOpenAnotherPPT_Click(object sender, RoutedEventArgs e)
        {
            OpenAnotherPPT();
        }

        private void GridConnectPresentation_Click(object sender, RoutedEventArgs e)
        {
            ConectPresentation();
        }

        public void ConectPresentation()
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

        // PowerPoint 열고 닫기
        private void OpenPPTRun_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Process.Start(VersionSelector.GetPowerPointPath());
            this.Close();
        }
    }
}
