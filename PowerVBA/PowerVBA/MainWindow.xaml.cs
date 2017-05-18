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
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using System.Diagnostics;
using System.IO;
using PowerVBA.Core.AvalonEdit.Replace;
using ICSharpCode.AvalonEdit.Folding;

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
        SQLiteConnector dbConnector;
        CodeInfo codeInfo;
        List<FileInfo> LibraryFiles = new List<FileInfo>();
        Stopwatch ParseSw = new Stopwatch();

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

            menuTabControl.SelectionChanged += MenuTabControl_SelectionChanged;
            codeTabControl.SelectionChanged += CodeTabControl_SelectionChanged;

            errorList.LineMoveRequest += ErrorList_LineMoveRequest;

            solutionExplorer.OpenRequest += SolutionExplorer_OpenRequest;
            solutionExplorer.OpenPropertyRequest += SolutionExplorer_OpenPropertyRequest;
            solutionExplorer.DeleteRequest += SolutionExplorer_DeleteRequest;

            backBtn.Click += BackBtn_Click;

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
                            catch
                            {
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

            RoutedCommand commAddItm = new RoutedCommand();
            commAddItm.InputGestures.Add(new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift));

            RoutedCommand commAddMethod = new RoutedCommand();
            commAddMethod.InputGestures.Add(new KeyGesture(Key.D, ModifierKeys.Control | ModifierKeys.Shift));

            RoutedCommand commAddVar = new RoutedCommand();
            commAddVar.InputGestures.Add(new KeyGesture(Key.V, ModifierKeys.Control | ModifierKeys.Shift));

            RoutedCommand commSaveAll = new RoutedCommand();
            commSaveAll.InputGestures.Add(new KeyGesture(Key.S, ModifierKeys.Control | ModifierKeys.Shift));


            var cb1 = new CommandBinding(commAddItm, Comm_ItmAdd);
            var cb2 = new CommandBinding(commAddMethod, Comm_MethodAdd);
            var cb3 = new CommandBinding(commAddVar, Comm_VarAdd);
            var cb4 = new CommandBinding(commSaveAll, Comm_SaveAll);

            this.CommandBindings.Add(cb1);
            this.CommandBindings.Add(cb2);
            this.CommandBindings.Add(cb3);
            this.CommandBindings.Add(cb4);

            #endregion

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

        private void Comm_SaveAll(object sender, ExecutedRoutedEventArgs e)
        {
            BtnAllFileSync_SimpleButtonClicked(sender);
        }

        private void Comm_VarAdd(object sender, ExecutedRoutedEventArgs e)
        {
            BtnAddVar_SimpleButtonClicked(sender);
        }

        private void Comm_MethodAdd(object sender, ExecutedRoutedEventArgs e)
        {
            BtnAddFunc_SimpleButtonClicked(sender);
        }

        private void Comm_ItmAdd(object sender, ExecutedRoutedEventArgs e)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Module);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void SolutionExplorer_DeleteRequest(object sender, VBComponentWrappingBase Data)
        {
            string Name = Data.CompName;
            if (connector.DeleteComponent(Data))
            {
                var itm = GetAllCodeTabs().Where(i => i.Header.ToString() == Name);
                codeTabControl.Items.Remove(itm.First());
            }
        }
        
        private void ErrorList_LineMoveRequest(Codes.TypeSystem.Error err)
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

        private void SolutionExplorer_OpenPropertyRequest()
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "솔루션 탐색기").ToList();
            if (tabItems.Count >= 1)
            {
                codeTabControl.SelectedItem = tabItems.First();
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


        private void SolutionExplorer_OpenRequest(object sender, VBComponentWrappingBase data)
        {
            if (data != null)
            {
                var itm = data.ToVBComponent2013();
                AddCodeTab(data);
            }
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

        private void Itm_DeleteRequest(object sender)
        {
            if (!dbConnector.RecentFile_Remove(((RecentFileListViewItem)sender).FileLocation))
                MessageBox.Show("알 수 없는 이유로 삭제에 실패했습니다.");
            else
                LVrecentFile.Items.Remove(sender);
        }

        private void Itm_CopyOpenRequest(object sender)
        {
            if (sender is RecentFileListViewItem itm)
            {
                string path = Path.GetFileNameWithoutExtension(itm.FileLocation);
                SetMessage($"'{path}' 프레젠테이션을 열고 있습니다.");
                InitalizeConnector(itm.FileLocation, true);
            }
        }

        private void Itm_OpenRequest(object sender)
        {
            RecentFileListViewItem itm = (RecentFileListViewItem)sender;
            SetMessage($"'{new FileInfo(itm.FileLocation).Name}' 프레젠테이션을 열고 있습니다.");
            InitalizeConnector(itm.FileLocation);
        }
        
        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            ParseSw.Restart();
            CodeSync(sender);
            btnUndo.IsEnabled = ((CodeEditor)sender).CanUndo;
            btnRedo.IsEnabled = ((CodeEditor)sender).CanRedo;
        }

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            programTabControl.SelectedIndex = 0;
            this.NoTitle = false;
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
                            e.Cancel = true;
                            return;
                    }
                }
            }

            AllClose();
        }

        public void AllClose()
        {
            connector?.Dispose();
            Environment.Exit(0);
        }

        private void MenuTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (menuTabControl.SelectedIndex == 0)
            {
                programTabControl.SelectedIndex = 2;
                menuTabControl.SelectedIndex = 1;

                if (programTabControl.SelectedIndex == 2) this.NoTitle = true;
            }
        }

        private void debugBtn_SimpleButtonClicked(object sender)
        {
            var itm = ((CodeTabEditor)((CloseableTabItem)codeTabControl.SelectedItem).Content).CodeEditor;
            itm.DeleteIndent();
        }

        private void FileTabControl_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (((TabItem)FileTabControl.SelectedItem).Header.ToString() == "닫기")
            {
                Environment.Exit(0);
            }
        }

        private void OpenPPTRun_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Process.Start(VersionSelector.GetPowerPointPath());
            this.Close();
        }

        private void BtnToPPT_SimpleButtonClicked(object sender)
        {
            connector.ActivateWindow();
        }

        private void AddRefBtn_SimpleButtonClicked(object sender)
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

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            var helperWindow = new HelperWindow();
            helperWindow.MoveHelpContext(((Button)sender).Tag.ToString());
            helperWindow.ShowDialog();
        }

        private void btnSync_MouseDown(object sender, MouseButtonEventArgs e)
        {
            runClass.Text = connector.GetClasses().Count.ToString();
        }
    }
}
