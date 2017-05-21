using PowerVBA.Controls.Customize;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.V2010.Connector;
using PowerVBA.V2013.Connector;
using PowerVBA.Wrap;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace PowerVBA
{

    // MainWindow Partial File :: Connect Part
    // MainWindow 부분 파일 :: 연결 부분


    public partial class MainWindow
    {

        private string ProjName { get; set; }


        /// <summary>
        /// 프로젝트 이름을 재설정합니다. customName이 비어있다면 현재의 프레젠테이션 연결기에서 이름을 가져옵니다.
        /// </summary>
        /// <param name="customName"></param>
        public void SetName(string customName = "")
        {
            if (customName != string.Empty)
            {
                ProjName = customName;
                this.Title = $"{customName} - PowerVBA";
            }
            else if (connector != null)
            {
                this.Title = connector.Name + " - PowerVBA";

                ProjName = connector.Name;

                if (connector.ReadOnly)
                    this.Title += " [읽기 전용]";
            }
        }

        /// <summary>
        /// 프레젠테이션 연결기를 초기화 합니다. PresentationWrappingBase를 인자로 받아 연결합니다.
        /// </summary>
        public void InitalizeConnector(PresentationWrappingBase pptWrapping)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                this.NoTitle = false;

                var version = VersionSelector.GetPPTVersion();
                if ((int)version < 14)
                {
                    MessageBox.Show("지원하지 않는 버전입니다.");
                    return;
                }

                PPTConnectorBase tmpConn = null;

                switch (version)
                {
                    case PPTVersion.PPT2010:
                        tmpConn = new PPTConnector2010(pptWrapping.ToPresentation2010());
                        break;
                    case PPTVersion.PPT2016:
                    case PPTVersion.PPT2013:
                        tmpConn = new PPTConnector2013(pptWrapping.ToPresentation2013());
                        break;
                }


                tmpConn.PresentationClosed += PPTCloseDetect;
                tmpConn.VBAComponentChange += ProjectFileChange;
                tmpConn.SlideChanged += SlideChangedDetect;
                tmpConn.ShapeChanged += ShapeChangedDetect;
                tmpConn.SelectionChanged += SelectionChangedDetect;
                tmpConn.PropertyChanged += TmpConn_PropertyChanged;

                connector = tmpConn;

                CodeSync(null);
                programTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }



        /// <summary>
        /// 프레젠테이션 연결기를 초기화 합니다. 파일 경로로 연결합니다.
        /// </summary>
        /// <param name="CopyOpen">복사해 열 것인지에 대한 여부를 나타냅니다.</param>
        public void InitalizeConnector(string FileLocation = "", bool CopyOpen = false)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                this.NoTitle = false;

                if (FileLocation != "" && !File.Exists(FileLocation))
                {
                    MessageBox.Show("파일 위치가 올바르지 않습니다.");
                    return;
                }
                var version = VersionSelector.GetPPTVersion();
                if ((int)version < 14)
                {
                    MessageBox.Show("지원하지 않는 버전입니다.");
                    return;
                }
                PPTConnectorBase tmpConn = null;

                switch (version)
                {
                    case PPTVersion.PPT2010:
                        if (FileLocation == string.Empty)
                            tmpConn = new PPTConnector2010();
                        else
                        {
                            
                            tmpConn = new PPTConnector2010(FileLocation, CopyOpen);
                        }
                        break;
                    case PPTVersion.PPT2016:
                    case PPTVersion.PPT2013:
                        if (FileLocation == string.Empty)
                            tmpConn = new PPTConnector2013();
                        else
                            tmpConn = new PPTConnector2013(FileLocation, CopyOpen);
                        break;
                }

                tmpConn.PresentationClosed += PPTCloseDetect;
                tmpConn.VBAComponentChange += ProjectFileChange;
                tmpConn.SlideChanged += SlideChangedDetect;
                tmpConn.ShapeChanged += ShapeChangedDetect;
                tmpConn.SelectionChanged += SelectionChangedDetect;
                tmpConn.PropertyChanged += TmpConn_PropertyChanged;

                connector = tmpConn;

                CodeSync(null);
                programTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

        private void TmpConn_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "AutoShapeUpdate":
                    projAnalyzer.AutoShapeUpdate = connector.AutoShapeUpdate;
                    break;
            }
        }



        private void PPTCloseDetect()
        {
            var result = MessageBox.Show("프레젠테이션이 PowerVBA의 코드가 저장되지 않은 상태에서 닫혔습니다.\r\n" +
                "코드 파일만 따로 저장하시겠습니까?\r\n" +
                "[예]를 누르시면 현재 열린 코드 파일만 추출되어 바탕화면에 저장됩니다.\r\n" +
                "[아니오]를 누르시면 코드 파일은 소멸되며 종료됩니다.", "PowerVBA", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    foreach (CloseableTabItem itm in codeTabControl.Items
                                                         .Cast<CloseableTabItem>()
                                                         .Where(i => i.Content.GetType() == typeof(CodeEditor)))
                    {
                        var editor = (CodeEditor)itm.Content;

                        string dirPath = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\{ProjName}\\";
                        string path = $"{dirPath}{itm.Header.ToString()}";

                        if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

                        File.Create(path).Dispose();

                        StreamWriter sw = new StreamWriter(path);

                        sw.Write(editor.Text);
                    }
                }
                catch (Exception)
                { }
            }
            Environment.Exit(0);
        }

        private void ProjectFileChange()
        {
            solutionExplorer.Update(connector);
        }


    }
}
