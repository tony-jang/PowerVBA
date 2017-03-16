using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerVBA.Windows;
using PowerVBA.Core.Connector;
using Microsoft.Win32;
using System.Windows.Threading;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Windows.AddWindows;
using static PowerVBA.Global.Globals;
using PowerVBA.V2013.Connector;
using PowerVBA.Core.Extension;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Wrap;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.Controls;

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : ChromeWindow
    {
        PPTConnectorBase Connector;
        public MainWindow()
        {
            InitializeComponent();
            
            this.Closing += MainWindow_Closing;
            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;
            CodeTabControls.SelectionChanged += CodeTabControls_SelectionChanged;

            MainDispatcher = Dispatcher;

            // Menu 초기화

            FileMenu.Items.Add(new MenuItem() { Header = "asdf" });


            RunVersion.Text = VersionSelector.GetPPTVersion().GetDescription();

            AddTab();
            AddTab();
            AddTab();
            AddTab();
        }
        #region [  코드 에디터(CodeEditor) 부분 코드  ]

        #endregion





        #region [  에디터 메뉴 탭  ]



        #region [  홈 탭 이벤트  ]

        private void BtnCopy_SimpleButtonClicked()
        {
            Clipboard.Clear();
            Clipboard.SetText(((CodeEditor)CodeTabControls.SelectedContent).SelectedText);
        }
        private void BtnPaste_SimpleButtonClicked()
        {
            if (Clipboard.ContainsText())
            {
                string t = Clipboard.GetText();
                CodeEditor editor = ((CodeEditor)CodeTabControls.SelectedContent);

                if (editor.SelectionLength != 0)
                {
                    editor.SelectedText = t;
                }
                else
                {
                    editor.TextArea.Document.Insert(editor.CaretOffset, t);
                }
                
            }

            
        }

        private void BtnUndo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControls.SelectedContent);
            if (editor.CanUndo) editor.Undo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        private void BtnRedo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControls.SelectedContent);
            if (editor.CanRedo) editor.Redo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        #endregion


        #region [  삽입 탭 이벤트  ]
        private void BtnAddClass_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, true);

            filewindow.ShowDialog();
        }

        private void BtnAddModule_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, false);

            filewindow.ShowDialog();
        }

        private void BtnAddSub_SimpleButtonClicked()
        {

        }

        private void BtnAddFunc_SimpleButtonClicked()
        {

        }

        private void BtnAddMouseOverTrigger_SimpleButtonClicked()
        {

        }

        private void BtnAddMouseClickTrigger_SimpleButtonClicked()
        {

        }

        private void BtnAddEnum_SimpleButtonClicked()
        {

        }

        private void BtnAddType_SimpleButtonClicked()
        {

        }
        #endregion


        #region [  슬라이드 탭 이벤트  ]
        private void BtnNewSlide_SimpleButtonClicked()
        {
            int SlideNumber = 0;


            switch (Connector.Version)
            {
                case PPTVersion.PPT2010:

                    break;
                case PPTVersion.PPT2013:
                    PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                    if (conn2013.Presentation.Slides.Count != 0) SlideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2013.Presentation.Slides.AddSlide(SlideNumber + 1, conn2013.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2013.Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
                    break;
                case PPTVersion.PPT2016:

                    break;
            }

            
        }

        private void BtnDelSlide_SimpleButtonClicked()
        {
            PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

            int SlideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            if (MessageBox.Show(SlideNumber + "슬라이드를 삭제합니다. 계속하시려면 예로 계속하세요.", "슬라이드 삭제 확인",
                MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                conn2013.Presentation.Slides[SlideNumber].Delete();
            }
        }
        #endregion

        #endregion


        #region [  윈도우 코드  ]

        public void AddTab()
        {
            CodeTabEditor codeTabEditor = new CodeTabEditor();

            CloseableTabItem codeTab = new CloseableTabItem()
            {
                Header = "ABC",
                Content = codeTabEditor
            };
            CodeTabControls.Items.Add(codeTab);
        }
        public void AddTab(VBComponentWrappingBase component)
        {
            switch (component.ClassVersion)
            {
                case PPTVersion.PPT2010:
                    break;
                case PPTVersion.PPT2013:
                    var comp2013 = component.ToVBComponent2013();

                    CodeTabEditor codeTabEditor = new CodeTabEditor();

                    CloseableTabItem codeTab = new CloseableTabItem()
                    {
                        Header = comp2013.Name,
                        Content = codeTabEditor
                    };
                    CodeTabControls.Items.Add(codeTab);
                    break;
                case PPTVersion.PPT2016:

                    break;
            }
            
        }

        public void RemoveTab()
        {

        }



        public void SetName(string customName = "")
        {
            if (customName!= "")
            {
                this.Title = $"{customName} - PowerVBA";
            }
            else if (Connector != null)
            {
                PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                this.Title = conn2013.Presentation.Name + " - PowerVBA";
                if (conn2013.Presentation.ReadOnly == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    this.Title += " [읽기 전용]";
                }
            }
        }




        #region [  이벤트  ]

        private void PPTCloseDetect()
        {
            var result = MessageBox.Show("프레젠테이션이 PowerVBA의 코드가 저장되지 않은 상태에서 닫혔습니다.\r\n코드는 저장한 상태에서 다시 여시겠습니까?\r\n이대로 닫게 되면 코드는 저장되지 않습니다.", "PowerVBA", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {

            }
            else
            {
                Environment.Exit(0);
            }

            
        }


        private void ProjectFileChange()
        {
            solutionExplorer.Update(Connector);
        }


        private void CodeTabControls_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            return;

            CodeEditor editor = ((CodeEditor)CodeTabControls.SelectedContent);


            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Connector?.Dispose();
        }
        ContextMenu FileMenu = new ContextMenu();
        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FileMenu.IsOpen = true;
            if (MainTabControl.SelectedIndex == 0) MainTabControl.SelectedIndex = 1;
        }
        #endregion

        #endregion




        #region [  초기 화면  ]
        private void GridOpenAnotherPpt_MouseLeave(object sender, MouseEventArgs e)
        {
            TBOpenAnotherPpt.FontStyle = FontStyles.Normal;
            while (TBOpenAnotherPpt.TextDecorations.Count != 0)
            {
                TBOpenAnotherPpt.TextDecorations.RemoveAt(0);
            }
        }

        private void GridOpenAnotherPpt_MouseMove(object sender, MouseEventArgs e)
        {
            TBOpenAnotherPpt.TextDecorations.Add(TextDecorations.Underline);
        }

        private void GridOpenAnotherPpt_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "프레젠테이션|*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm"
            };
            if (ofd.ShowDialog().Value)
            {
                tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

                InitalizeConnector(ofd.FileName);
            }
            
        }

        private void BtnNewPPT_Click(object sender, RoutedEventArgs e)
        {

            PPTVersion ver = VersionSelector.GetPPTVersion();

            if (ver != PPTVersion.PPT2013 && ver != PPTVersion.PPT2010)
            {
                MessageBox.Show($"죄송합니다. {ver.ToString()}는 지원하지 않는 버전입니다.");
                return;
            }

            tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

            InitalizeConnector();

        }

        public void InitalizeConnector(string FileLocation = "")
        {
            Dispatcher.Invoke(new Action(() =>
            {
                this.NoTitle = false;
                PPTConnectorBase tmpConn;
                if (FileLocation == "")
                {
                    tmpConn = new PPTConnector2013();
                }
                else
                {
                    tmpConn = new PPTConnector2013(FileLocation);
                }

                tmpConn.PresentationClosed += PPTCloseDetect;
                tmpConn.VBAComponentChange += ProjectFileChange;
                tmpConn.SlideChanged += SlideChangedDetect;
                tmpConn.ShapeChanged += ShapeChangedDetect;
                tmpConn.SectionChanged += SectionChangedDetect;


                Connector = tmpConn;

                PropGrid.SelectedObject = Connector.ToConnector2013().Presentation;

                ProgramTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

        private void SectionChangedDetect()
        {
            projAnalyzer.SectionCount = Connector.ToConnector2013().Presentation.SectionCount;
        }

        private void ShapeChangedDetect()
        {
            int count = 0;
            Connector.ToConnector2013().Presentation.Slides.Cast<Microsoft.Office.Interop.PowerPoint.Slide>().ToList().ForEach((i) => count += i.Shapes.Count);
            projAnalyzer.ShapeCount = count;
        }

        private void SlideChangedDetect()
        {
            projAnalyzer.SlideCount = Connector.ToConnector2013().Presentation.Slides.Count;
        }

        private void BtnNewAssistPPT_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnNewVirtualPPT_Click(object sender, RoutedEventArgs e)
        {
            this.NoTitle = false;
            ProgramTabControl.SelectedIndex = 0;
            SetName("가상 프레젠테이션 1");
        }
        #endregion
        


    }
}
