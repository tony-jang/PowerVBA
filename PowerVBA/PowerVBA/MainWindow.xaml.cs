using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerVBA.Windows;
using PowerVBA.Core.Connector;
using System.Windows.Threading;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Windows.AddWindows;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Wrap;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.Controls;
using System.Threading;
using System.ComponentModel;
using System.Collections.Generic;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using System.Diagnostics;
using System.IO;
using PowerVBA.Core.AvalonEdit.Replace;
using ICSharpCode.AvalonEdit.Folding;
using Microsoft.Win32;
using static PowerVBA.Global.Globals;
using PowerVBA.Codes.Extension;
using ppt = Microsoft.Office.Interop.PowerPoint;
using PowerVBA.Core.Extension;
using PowerVBA.Resources;

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : ChromeWindow
    {
        PPTConnectorBase connector;
        StartupWindow suWindow = new StartupWindow();
        BackgroundWorker bg;
        Thread loadThread;
        SQLiteConnector dbConnector = new SQLiteConnector();
        CodeInfo codeInfo;
        List<FileInfo> libraryFiles = new List<FileInfo>();
        Stopwatch ParseSw = new Stopwatch();

        /// <summary>
        /// 연결되어 있는지에 대한 여부를 확인합니다.
        /// </summary>
        bool IsConnected = false;

        public MainWindow()
        {
            InitializeComponent();

            codeInfo = new CodeInfo();
            new VBAParser(codeInfo);

            #region [  BackgroundWorker  ]

            bg = new BackgroundWorker();
            bg.DoWork += bg_DoWork;
            bg.RunWorkerCompleted += bg_RunWorkerCompleted;

            bg.WorkerReportsProgress = true;
            #endregion

            #region [  EventHandler  ]

            this.Closing += MainWindow_Closing;
            this.Loaded += MainWindow_Loaded;
            this.Activated += MainWindow_Activated;

            codeTabControl.SelectionChanged += CodeTabControl_SelectionChanged;

            errorList.LineMove += ErrorList_LineMove;

            solutionExplorer.Open += SolutionExplorer_Open;
            solutionExplorer.Delete += SolutionExplorer_Delete;

            solutionExplorer.OpenProperty += SolutionExplorer_OpenProperty;
            solutionExplorer.OpenObjectBrowser += SolutionExplorer_OpenBrowser;
            solutionExplorer.OpenShapeExplorer += SolutionExplorer_OpenShapeExplorer;
            

            projAnalyzer.ShapeSyncRequest += ProjAnalyzer_ShapeSyncRequest;

            backBtn.Click += backBtn_Click;

            #endregion

            Thread parseThr = new Thread(() =>
            {
                while (true)
                {
                    if (ParseSw.ElapsedMilliseconds > 500)
                    {
                        Stopwatch debugSw = new Stopwatch();
                        debugSw.Start();

                        #region [  Parsing Process  ]

                        List<(string, string)> CodeLists = new List<(string, string)>();

                        Dispatcher.Invoke(new Action(() => {
                            CodeLists = codeTabControl.Items
                                                      .Cast<TabItem>()
                                                      .Where(t => t.Content.GetType() == typeof(CodeEditor))
                                                      .Select(t => (((CodeEditor)t.Content).Text, t.Header.ToString())).ToList();
                        }));

                        new VBAParser(codeInfo).Parse(CodeLists);

                        errorList.Dispatcher.Invoke(new Action(() =>
                        {
                            SetMessage("코드 파싱에 소요된 시간 : " + debugSw.ElapsedMilliseconds.ToString() + "ms"); // DEBUG:
                            errorList.SetError(codeInfo.ErrorList);
                        }));

                        Dispatcher.Invoke(new Action(() => {
                            try
                            {
                                foreach (CloseableTabItem itm in codeTabControl.Items)
                                {
                                    if (itm.Content.GetType() == typeof(CodeEditor))
                                    {
                                        CodeEditor editor = itm.Content as CodeEditor;
                                        var list = codeInfo.Lines.Where(i => i.FileName == itm.Header.ToString());

                                        if (list.Count() >= 1)
                                        {
                                            var nFolding = list.First().Foldings
                                                .Select(i =>
                                                new NewFolding(editor.Document.GetOffset(i.StartInt, 0), editor.Document.GetLineByNumber(i.EndInt).EndOffset)
                                                {
                                                    Name = editor.Text.Substring(editor.Document.GetOffset(i.StartInt, 0),
                                                                                 editor.Document.GetLineByNumber(i.StartInt).Length)
                                                });

                                            editor.foldingManager.UpdateFoldings(nFolding, -1);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                                SetMessage("코드 폴딩 처리중에 오류가 발생한 것 같습니다.");
                            }
                        }));

                        ParseSw.Reset();
                        #endregion
                    }
                    Thread.Sleep(10);
                }
            });

            parseThr.SetApartmentState(ApartmentState.STA);
            parseThr.Start();

            #region [  Add Command Binding  ]

            AddCommandBinding(new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift), Comm_ItmAdd);
            AddCommandBinding(new KeyGesture(Key.D, ModifierKeys.Control | ModifierKeys.Shift), Comm_MethodAdd);
            AddCommandBinding(new KeyGesture(Key.V, ModifierKeys.Control | ModifierKeys.Shift), Comm_VarAdd);
            AddCommandBinding(new KeyGesture(Key.S, ModifierKeys.Control | ModifierKeys.Shift), Comm_SaveAll);
            AddCommandBinding(new KeyGesture(Key.O, ModifierKeys.Control), Comm_Open);

            #endregion
        }

        private void BtnRefresh_ButtonClick(object sender)
        {
            RefreshPPTItem();
        }

        public void RefreshPPTItem()
        {
            lvOpenPPTs.Items.Clear();

            var itm = new ppt.Application();

            foreach (ppt.Presentation presentation in itm.Presentations)
            {

                var imgBtn = new ImageButton()
                {
                    TextAlignment = TextAlignment.Left,
                    ButtonMode = ImageButton.ButtonModes.LongWidth,
                    BackImage = ResourceImage.GetIconImage("PPTIcon"),
                    Content = presentation.Name + (((Bool2)presentation.ReadOnly) ? " [읽기 전용]" : string.Empty),
                    Tag = presentation
                };

                imgBtn.ButtonClick += ImgBtn_ButtonClick;

                lvOpenPPTs.Items.Add(imgBtn);
            }            
        }

        private void ImgBtn_ButtonClick(object sender)
        {
            InitalizeConnector(new V2013.WrapClass.PresentationWrapping(((ppt.Presentation)((Control)sender).Tag)));
        }

        public void AddCommandBinding(KeyGesture gesture, ExecutedRoutedEventHandler commandEvent)
        {
            RoutedCommand rCommand = new RoutedCommand();

            rCommand.InputGestures.Add(gesture);

            this.CommandBindings.Add(new CommandBinding(rCommand, commandEvent));
        }



        bool LoadComplete = false;
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Opacity = 0.0;
            loadThread = new Thread(() =>
            {
                suWindow = new StartupWindow();

                Dispatcher.FromThread(loadThread).Invoke(new Action(() =>
                {
                    suWindow.Show();
                }), DispatcherPriority.Background);

                while (!LoadComplete) { }
                suWindow.Close();
            });

            loadThread.SetApartmentState(ApartmentState.STA);
            Dispatcher.Invoke(new Action(() =>
            {
                loadThread.Start();
            }), DispatcherPriority.Background);

            bg.RunWorkerAsync();
        }

        #region [  BackgroundWorker  ]

        private void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Opacity = 1.0;
            LoadComplete = true;

            this.Activate();
        }

        private void bg_DoWork(object sender, DoWorkEventArgs e)
        {
            MainDispatcher = Dispatcher;

            RecentFileSet();
            RecentFolderSet();
        }

        #endregion

        private void ProjAnalyzer_ShapeSyncRequest()
        {
            if (IsConnected) ShapeChangedDetect();
        }

        private void SolutionExplorer_OpenShapeExplorer()
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "도형 탐색기").ToList();
            if (tabItems.Count >= 1)
            {
                codeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "도형 탐색기",
                    Content = new ShapeExplorer(connector.Slide, connector)
                };
                codeTabControl.Items.Add(itm);
                codeTabControl.SelectedItem = itm;
            }
        }

        private void SolutionExplorer_OpenBrowser()
        {
            
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
            if (connector == null) return;
            
            Dispatcher.Invoke(() =>
            {
                try
                {
                    var codeTabs = GetAllCodeTabs();


                    

                    foreach (VBComponentWrappingBase item in connector.GetFiles())
                    {
                        var itm = codeTabs.Where(i => i.Header.ToString() == item.CompName).FirstOrDefault();
                        if (itm != null)
                        {
                            CodeEditor editor = (CodeEditor)itm.Content;
                            if (editor.Text != item.Code && editor.LastText != item.Code)
                            {
                                if (editor.Saved)
                                {
                                    editor.Text = item.Code;
                                }
                                else
                                {
                                    var msg = MessageBox.Show("파워포인트의 코드와 내 파일의 저장되지 않은 내용이 존재합니다. 변경하시겠습니까?" +
                                                 "\r\n파워포인트의 코드로 변환하려면 [예]를 누르세요." +
                                                 "\r\n현재 내 파일의 코드로 변경하려면 [아니오]를 누르세요.", "코드 충돌", MessageBoxButton.YesNo);

                                    if (msg == MessageBoxResult.Yes)
                                    {
                                        editor.Text = item.Code;
                                    }
                                }
                                editor.Save();
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
            });   
        }

        // 커멘드 - 열기
        private void Comm_Open(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                SetProgramTab(ProgramTabMenus.FileTab);
                SetFileTabMenu(FileTabMenus.Open);
            }
        }


        // 커멘드 - 전체 저장
        private void Comm_SaveAll(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                BtnAllFileSync_ButtonClick(sender);
            }
        }

        // 커멘드 - 변수 추가
        private void Comm_VarAdd(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                BtnAddVar_ButtonClick(sender);
            }
        }

        // 커멘드 - 메소드 추가
        private void Comm_MethodAdd(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                BtnAddFunc_ButtonClick(sender);
            }
        }

        // 커멘드 - 아이템 추가
        private void Comm_ItmAdd(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Module);

                SolutionExplorer_Open(this, filewindow.ShowDialog());
            }
        }

        private void SolutionExplorer_Delete(object sender, VBComponentWrappingBase Data)
        {
            string Name = Data.CompName;
            if (connector.DeleteComponent(Data))
            {
                var itm = GetAllCodeTabs().Where(i => i.Header.ToString() == Name);
                codeTabControl.Items.Remove(itm.First());
            }
        }
        
        private void ErrorList_LineMove(Codes.TypeSystem.Error err)
        {
            try
            {
                CloseableTabItem tabItm = codeTabControl.Items.Cast<CloseableTabItem>()
                                                            .Where(i => i.Header.ToString() == err.FileName).FirstOrDefault();

                CodeEditor codeEditor = (CodeEditor)tabItm.Content;

                if (codeEditor == null || tabItm == null) return;

                codeEditor.ScrollToLine(err.Region.Begin.Line);
                codeEditor.SelectionLength = 0;
                codeEditor.CaretOffset = codeEditor.Document.GetOffset(err.Region.Begin);
                codeEditor.SelectionLength = codeEditor.Document.GetLineByOffset(codeEditor.SelectionStart).Length;

                codeTabControl.SelectedItem = tabItm;
                codeEditor.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());   
            }   
        }


        private void SolutionExplorer_OpenProperty()
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "솔루션 탐색기").ToList();
            if (tabItems.Count >= 1)
            {
                var itm = tabItems.First();
                codeTabControl.SelectedItem = itm;
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "솔루션 탐색기",
                    Content = new ProjectProperty(connector)
                };
                codeTabControl.Items.Add(itm);
                codeTabControl.SelectedItem = itm;
            }
        }


        private void SolutionExplorer_Open(object sender, VBComponentWrappingBase data)
        {
            if (data != null)
            {
                var itm = data.ToVBComponent2013();
                AddCodeTab(data);
            }
        }


        // 아이템 삭제를 요청합니다.
        private void Itm_DeleteRequest(object sender)
        {
            if (!dbConnector.RecentFileRemove(((RecentFileListViewItem)sender).FileLocation))
                MessageBox.Show("알 수 없는 이유로 삭제에 실패했습니다.");
            else
                lvRecentFile.Items.Remove(sender);
        }

        // 아이템 복사를 요청합니다.
        private void Itm_CopyOpenRequest(object sender)
        {
            if (sender is RecentFileListViewItem itm)
            {
                string path = Path.GetFileNameWithoutExtension(itm.FileLocation);
                SetMessage($"'{path}' 프레젠테이션을 열고 있습니다.");
                InitalizeConnector(itm.FileLocation, true);
            }
        }
        
        // 열기 요청되었을때
        private void Itm_OpenRequest(object sender)
        {
            RecentFileListViewItem itm = (RecentFileListViewItem)sender;
            SetMessage($"'{new FileInfo(itm.FileLocation).Name}' 프레젠테이션을 열고 있습니다.");
            InitalizeConnector(itm.FileLocation);
        }
        
        // CodeEditor의 코드가 변경되었을때
        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            ParseSw.Restart();
            CodeSync(sender);
            btnUndo.IsEnabled = ((CodeEditor)sender).CanUndo;
            btnRedo.IsEnabled = ((CodeEditor)sender).CanRedo;
        }

        // 파일에서 뒤로 돌아가기 버튼 누를때
        private void backBtn_Click(object sender, RoutedEventArgs e)
        {
            SetProgramTab(ProgramTabMenus.MainEdit);
        }

        private void CodeTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FindReplaceDialog.CloseWindow();
            CloseableTabItem tabItm = ((CloseableTabItem)codeTabControl.SelectedItem);

            bool flag, triggerFlag;

            if (tabItm == null || tabItm.Content.GetType() != typeof(CodeEditor))
            {
                flag = false;
                triggerFlag = false;

                btnUndo.IsEnabled = false;
                btnRedo.IsEnabled = false;
            }
            else
            {
                var editor = (CodeEditor)tabItm.Content;

                btnUndo.IsEnabled = editor.CanUndo;
                btnRedo.IsEnabled = editor.CanRedo;
                
                flag = true;

                triggerFlag = tabItm.Header.ToString().EndsWith(".bas"); // 모듈의 경우에만 추가 버튼 활성화
            }

            btnAddSub.IsEnabled = flag;
            btnAddFunc.IsEnabled = flag;
            btnAddProp.IsEnabled = flag;
            btnAddVar.IsEnabled = flag;
            btnAddConst.IsEnabled = flag;
            btnFileSync.IsEnabled = flag;

            btnAddMouseOverTrigger.IsEnabled = triggerFlag;
            btnAddMouseClickTrigger.IsEnabled = triggerFlag;
        }
        

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (CloseRequest()) AllClose();
            else e.Cancel = true;
        }

        /// <summary>
        /// 닫기 요청을 합니다. 예 또는 아니오가 눌리면 true(닫기 요청)를(을) 취소가 눌리면 false(닫기 취소)를 반환합니다. 
        /// </summary>
        /// <returns></returns>
        public bool CloseRequest()
        {
            if (connector != null)
            {
                var nonSavedFile = GetAllCodeTabs()
                                      .Select(i => (CodeEditor)i.Content)
                                      .Where(i => i.Saved);

                if (!connector.Saved || nonSavedFile.Count() != 0)
                {
                    var result = MessageBox.Show(connector.Name + "에 저장되지 않은 내용이 있습니다. 저장하시겠습니까?", "저장되지 않은 내용", MessageBoxButton.YesNoCancel);

                    switch (result)
                    {
                        case MessageBoxResult.Yes: // 저장후 닫기 요청

                            nonSavedFile.ToList().ForEach(i => i.Save());

                            if (!connector.Save()) // 저장
                            {
                                MessageBox.Show("저장에 실패했습니다. [읽기 전용]이거나 폰트가 없을 수 있습니다. 다른 이름으로 저장하기를 해주세요.", "저장 실패");
                                goto case MessageBoxResult.Cancel;
                            }
                            break;
                        case MessageBoxResult.Cancel: // 닫기 취소
                            return false;
                    }
                }
            }
            return true;
            
        }

        // 모든 커넥터와의 연결을 끊고 프로그램을 종료합니다.
        public void AllClose()
        {
            connector?.Dispose();
            Environment.Exit(0);
        }

        private void debugBtn_ButtonClick(object sender)
        {
            var itm = ((CodeTabEditor)((CloseableTabItem)codeTabControl.SelectedItem).Content).CodeEditor;
            itm.DeleteIndent();
        }


        private void TiSaveItem_MouseDown(object sender, MouseButtonEventArgs e)
        {
            BtnAllFileSync_ButtonClick(sender);

            var itm = GetAllCodeEditors();
            itm.ForEach(editor => editor.Save());
            connector.Save();

            SetProgramTab(ProgramTabMenus.MainEdit);

            e.Handled = true;
        }

        private void fileTabClose_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (CloseRequest())
            {
                AllClose();
            }
            
            e.Handled = true;
        }


        // PowerPoint 열고 닫기
        private void OpenPPTRun_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Process.Start(VersionSelector.GetPowerPointPath());
            this.Close();
        }

        
        private void BtnToPPT_ButtonClick(object sender)
        {
            connector.ActivateWindow();
        }


        private void AddRefBtn_ButtonClick(object sender)
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "참조 관리자").ToList();
            if (tabItems.Count >= 1)
            {
                codeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "참조 관리자",
                    Content = new ReferenceManager(connector)
                };
                codeTabControl.Items.Add(itm);
                codeTabControl.SelectedItem = itm;
            }
        }


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
                OpenAnotherPPT();
            }
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation"))
            {
                ConectPresentation();
            }
        }

        public void OpenAnotherPPT()
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "프레젠테이션|*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm"
            };
            if (ofd.ShowDialog().Value)
            {
                tbProcessInfoTB.Text = "프레젠테이션을 열고 있습니다.";

                dbConnector.RecentFileAdd(ofd.FileName);

                InitalizeAll();
                InitalizeConnector(ofd.FileName);
            }
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

        private void btnNewPPT_Click(object sender, RoutedEventArgs e)
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
        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            var helperWindow = new HelperWindow();
            helperWindow.MoveHelpContext(((Button)sender).Tag.ToString());
            helperWindow.ShowDialog();
        }

        // 동기화 버튼
        private void btnSync_MouseDown(object sender, MouseButtonEventArgs e)
        {
            runClass.Text = connector.GetClasses().Count.ToString();
        }

        private void btnSaveAsSetFile_ButtonClick(object sender)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "매크로 프레젠테이션 파일|*.pptm";

            if (sfd.ShowDialog().Value)
            {
                dbConnector.RecentFolderAdd(sfd.FileName);

                var itm = GetAllCodeEditors();
                itm.ForEach(editor => editor.Save());

                if (connector.SaveAs(sfd.FileName))
                {
                    NameSet();
                }
            }
        }


        void SetProgramTab(ProgramTabMenus programTabMenu)
        {

            switch (programTabMenu)
            {
                case ProgramTabMenus.Startup:
                case ProgramTabMenus.MainEdit:
                    this.NoTitle = false;
                    break;
                case ProgramTabMenus.FileTab:
                    this.NoTitle = true;
                    break;
                default:
                    this.NoTitle = false;
                    break;
            }

            programTabControl.SelectedIndex = (int)programTabMenu;
        }

        enum ProgramTabMenus
        {
            Startup = 1,
            MainEdit = 0,
            FileTab = 2,
        }

        void SetMainMenuTab(MainTabMenus mainTabMenu)
        {
            menuTabControl.SelectedIndex = (int)mainTabMenu;
        }

        enum MainTabMenus
        {
            Home = 1,
            Insert = 2,
            Project = 3,
        }

        void SetFileTabMenu(FileTabMenus fileTabMenu)
        {
            fileTabControl.SelectedIndex = (int)fileTabMenu;
        }
        
        enum FileTabMenus
        {
            Information = 0,
            NewPresentation = 1,
            Open = 2,
            Save = 3,
            SaveAs = 4,
            Close = 5,
        }

        private void btnFileTabOpen_ButtonClick(object sender)
        {
            OpenAnotherPPT();
        }

        private void FileTab_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            SetProgramTab(ProgramTabMenus.FileTab);

            e.Handled = true;
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tc = (TabControl)sender;

            if (tc.SelectedItem is TabItem ti)
            {
                if (ti.Tag?.ToString() == "connPresentation")
                {
                    RefreshPPTItem();
                }
            }

        }

        private void btnSync_ButtonClick(object sender)
        {
            runFunc.Text = codeInfo
        }
    }
}
