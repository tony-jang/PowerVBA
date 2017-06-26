using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

using PowerVBA.Core.AvalonEdit;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.AvalonEdit.Replace;
using PowerVBA.Core.DataBase;
using PowerVBA.Core.Connector;
using PowerVBA.Controls.Customize;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using PowerVBA.Codes.Parsing;
using PowerVBA.Enums;
using PowerVBA.Windows.AddWindows;
using PowerVBA.Wrap;
using PowerVBA.Windows;

using ICSharpCode.AvalonEdit.Folding;

using static PowerVBA.Global.Globals;
using System.Text;

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
        List<string> parseFiles = new List<string>();
        Stopwatch parseSw = new Stopwatch();

        List<VBComponentWrappingBase> ComponentFiles;

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

            outline.LinePointMove += Outline_LinePointMove;

            backBtn.Click += BackBtn_Click;

            #endregion

            Thread parseThr = new Thread(() =>
            {
                while (true)
                {
                    if (parseSw.ElapsedMilliseconds > 500)
                    {
                        Stopwatch debugSw = new Stopwatch();
                        debugSw.Start();

                        #region [  Parsing Process  ]

                        List<(string, string)> CodeLists = new List<(string, string)>();

                        Dispatcher.Invoke(new Action(() =>
                        {
                            CodeLists = codeTabControl.Items
                                .Cast<TabItem>()
                                .Where(t => parseFiles.Contains(t.Header.ToString()) && t.Content.GetType() == typeof(CodeEditor))
                                .Select(t => (t.Header.ToString(), ((CodeEditor)t.Content).Text)).ToList();
                        }));


                        new VBAParser(codeInfo).Parse(CodeLists);

                        errorList.Dispatcher.Invoke(new Action(() =>
                        {
#if DEBUG
                            SetMessage("코드 파싱에 소요된 시간 : " + debugSw.ElapsedMilliseconds.ToString() + "ms"); // DEBUG:
#endif
                            errorList.SetError(codeInfo.ErrorList);
                            //MessageBox.Show(string.Join(", ", codeInfo.Variables.Select(i => i.Name)));
                        }));

                        Dispatcher.Invoke(new Action(() => {
                            try
                            {
                                outline.ClearFile(codeInfo.CodeFiles.Select(i => i.FileName).ToArray());
                                foreach(var itm in codeInfo.CodeFiles)
                                {
                                    foreach(var var in itm.Variables)
                                    {
                                        outline.AddVariable(var, itm.FileName);
                                    }
                                    foreach(var func in itm.Functions)
                                    {
                                        outline.AddFunction(func, itm.FileName);
                                    }
                                    foreach(var sub in itm.Subs)
                                    {
                                        outline.AddSub(sub, itm.FileName);
                                    }
                                    foreach (var enumitm in itm.Enums)
                                    {
                                        outline.AddEnum(enumitm, itm.FileName);
                                    }
                                }

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

                        parseSw.Reset();
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
            AddCommandBinding(new KeyGesture(Key.F1), Comm_Help);
            AddCommandBinding(new KeyGesture(Key.O, ModifierKeys.Control), Comm_Open);
            AddCommandBinding(new KeyGesture(Key.D, ModifierKeys.Control), Comm_Debug);

            #endregion
        }

        private void Outline_LinePointMove(object sender, LinePointEventArgs e)
        {
            try
            {
                var tabItm = GetAllCodeTabs()
                    .Where(i => i.Header.ToString() == e.FileName).FirstOrDefault();

                CodeEditor codeEditor = (CodeEditor)tabItm.Content;

                

                if (codeEditor == null) return;

                codeEditor.ScrollToLine(e.LinePoint.Line);
                codeEditor.SelectionLength = 0;
                codeEditor.CaretOffset = codeEditor.Document.GetOffset(e.LinePoint.Line, e.LinePoint.Offset);
                codeEditor.SelectionLength = e.LinePoint.Length;

                codeTabControl.SelectedItem = tabItm;
                codeEditor.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Comm_Debug(object sender, ExecutedRoutedEventArgs e)
        {
            MessageBox.Show("Debug Action Activate");

            //connector.ToConnector2013().Presentation.VBProject.

            var itm = connector.GetFiles().Select(i => i.File);

            foreach(var codeFile in itm)
            {
                MessageBox.Show(codeFile.FileName);
            }

            return;

            connector.ToConnector2013().VBProject.Name = "Test";
            connector.ToConnector2013().VBProject.Description = "This Project Is Test Project";
        }

        private void Comm_Help(object sender, ExecutedRoutedEventArgs e)
        {
            HelperWindow hw = new HelperWindow();

            hw.ShowDialog();
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
            Dispatcher.Invoke(() => { RefreshOpenFolder(); });
        }

        #endregion

        private void ProjAnalyzer_ShapeSyncRequest()
        {
            if (IsConnected) ShapeChangedDetect();
        }

        private void SolutionExplorer_OpenBrowser()
        {
            
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
            if (connector == null)
                return;
            
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
                                    var msg = MessageBox.Show($"[{item.CompName}] 파일을 읽던 중에\r\n" +
                                        "파워포인트의 코드와 내 파일의 저장되지 않은 내용을 발견했습니다. 변경하시겠습니까?" +
                                        "\r\n파워포인트의 코드로 변환하려면 [예]를 누르세요." +
                                        "\r\n현재 내 파일의 코드로 변경하려면 [아니오]를 누르세요.", "코드 충돌", MessageBoxButton.YesNo);

                                    if (msg == MessageBoxResult.Yes)
                                        editor.Text = item.Code;
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

        #region [  Commands  ]

        // 커멘드 - 열기
        private void Comm_Open(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
            {
                SetProgramTab(ProgramTabMenu.FileTab);
                SetFileTabMenu(FileTabMenu.Open);
            }
        }
        
        // 커멘드 - 전체 저장
        private void Comm_SaveAll(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
                BtnAllFileSync_ButtonClick(sender);
        }

        // 커멘드 - 변수 추가
        private void Comm_VarAdd(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
                BtnAddVar_ButtonClick(sender);
        }

        // 커멘드 - 메소드 추가
        private void Comm_MethodAdd(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsConnected)
                BtnAddFunc_ButtonClick(sender);
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

        #endregion
        
        private void SolutionExplorer_Delete(object sender, VBComponentWrappingBase Data)
        {
            string Name = Data.CompName;
            if (connector.DeleteComponent(Data))
            {
                var itm = GetAllCodeTabs().Where(i => i.Header.ToString() == Name);
                if (itm.Count() != 0)
                    codeTabControl.Items.Remove(itm.First());
            }
        }
        
        private void ErrorList_LineMove(Error err)
        {
            try
            {
                CloseableTabItem tabItm = codeTabControl.Items.Cast<CloseableTabItem>()
                    .Where(i => i.Header.ToString() == err.FileName).FirstOrDefault();

                CodeEditor codeEditor = (CodeEditor)tabItm.Content;

                if (codeEditor == null || tabItm == null) return;

                codeEditor.ScrollToLine(err.Line);
                codeEditor.SelectionLength = 0;
                codeEditor.CaretOffset = codeEditor.Document.GetOffset(err.Line, 0);
                codeEditor.SelectionLength = codeEditor.Document.GetLineByOffset(codeEditor.SelectionStart).Length;

                codeTabControl.SelectedItem = tabItm;
                codeEditor.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());   
            }   
        }

        private void SolutionExplorer_OpenShapeExplorer()
        {
            AddTabItem(new ShapeExplorer(connector), "도형 탐색기");
        }

        private void SolutionExplorer_OpenProperty()
        {
            bool err = false;
            var prop = new ProjectProperty(connector, ref err);

            if (!err)
            {
                AddTabItem(prop, "프로젝트 속성").SaveCloseRequest += prop.SavecloseRequest;
            }
            else
            {
                MessageBox.Show("프로젝트 속성을 열던 중 오류가 발생했습니다.");
            }
        }

        private void AddRefBtn_ButtonClick(object sender)
        {
            AddTabItem(new ReferenceManager(connector), "참조 관리자");
        }

        public CloseableTabItem AddTabItem(object Item, string str)
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == str).ToList();
            if (tabItems.Count >= 1)
            {
                var itm = tabItems.First();
                codeTabControl.SelectedItem = itm;

                return itm;
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = str,
                    Content = Item
                };
                codeTabControl.Items.Add(itm);
                codeTabControl.SelectedItem = itm;

                return itm;
            }
        }
        
        private void SolutionExplorer_Open(object sender, VBComponentWrappingBase data)
        {
            if (data != null)
            {
                codeInfo.AddFile(data.File);

                var itm = data.ToVBComponent2013();
                AddCodeTab(data);
            }
        }

        #region [  Presentation Item  ]

        // 아이템 삭제를 요청합니다.
        private void Itm_DeleteRequest(object sender)
        {
            if (!dbConnector.FileTable.Remove(((RecentFileListViewItem)sender).FileLocation))
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

        #endregion

        // CodeEditor의 코드가 변경되었을때
        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            CodeEditor editor = (CodeEditor)sender;
            parseFiles.Add(editor.ConnectedFile.CompName);
            parseSw.Restart();
            btnUndo.IsEnabled = editor.CanUndo;
            btnRedo.IsEnabled = editor.CanRedo;
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

                gridCaretInfo.Visibility = Visibility.Hidden;
            }
            else
            {
                var editor = (CodeEditor)tabItm.Content;

                btnUndo.IsEnabled = editor.CanUndo;
                btnRedo.IsEnabled = editor.CanRedo;
                
                flag = true;

                triggerFlag = tabItm.Header.ToString().EndsWith(".bas"); // 모듈의 경우에만 추가 버튼 활성화
                
                gridCaretInfo.Visibility = Visibility.Visible;
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
            if (connector == null) return true;

            var nonSavedFile = GetAllCodeTabs()
                .Select(i => (CodeEditor)i.Content)
                .Where(i => i.Saved);

            if (!connector.Saved || !CodeSaved)
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
            return true;
        }

        // 모든 커넥터와의 연결을 끊고 프로그램을 종료합니다.
        public void AllClose()
        {
            connector?.Dispose();
            Environment.Exit(0);
        }
        
        private void TiSaveItem_MouseDown(object sender, MouseButtonEventArgs e)
        {
            BtnAllFileSync_ButtonClick(sender);

            var itm = GetAllCodeEditors();
            itm.ForEach(editor => editor.Save());
            connector.Save();

            SetProgramTab(ProgramTabMenu.MainEdit);

            e.Handled = true;
        }
        
        private void BtnToPPT_ButtonClick(object sender)
        {
            connector.ActivateWindow();
        }
        
        private void Buttons_MouseLeave(object sender, MouseEventArgs e)
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

        // 동기화 버튼
        private void BtnInfoSync_MouseDown(object sender, MouseButtonEventArgs e)
        {
            runClass.Text = connector.GetClasses().Count.ToString();
        }
                
        void SetProgramTab(ProgramTabMenu programTabMenu)
        {

            switch (programTabMenu)
            {
                case ProgramTabMenu.Startup:
                case ProgramTabMenu.MainEdit:
                    this.NoTitle = false;
                    break;
                case ProgramTabMenu.FileTab:
                    this.NoTitle = true;
                    break;
                default:
                    this.NoTitle = false;
                    break;
            }

            programTabControl.SelectedIndex = (int)programTabMenu;
        }

        void SetMainMenuTab(MainTabMenu mainTabMenu)
        {
            menuTabControl.SelectedIndex = (int)mainTabMenu;
        }

        void SetFileTabMenu(FileTabMenu fileTabMenu)
        {
            fileTabControl.SelectedIndex = (int)fileTabMenu;
        }

        // 찾아보기 버튼 클릭해서 '열기'
        private void BtnFileTabOpen_ButtonClick(object sender)
        {
            OpenAnotherPPT();
        }

        private void FileTab_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            SetProgramTab(ProgramTabMenu.FileTab);

            e.Handled = true;
        }

    }
}
