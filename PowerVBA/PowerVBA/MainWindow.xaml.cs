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
using System.Threading;
using System.ComponentModel;
using System.Collections.Generic;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using System.Diagnostics;
using System.IO;
using ICSharpCode.AvalonEdit.Folding;
using PowerVBA.V2010.Connector;
using PowerVBA.Core.AvalonEdit.Replace;

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
            
            bg = new BackgroundWorker();
            bg.DoWork += bg_DoWork;
            bg.RunWorkerCompleted += bg_RunWorkerCompleted;

            bg.WorkerReportsProgress = true;

            this.Closing += MainWindow_Closing;
            this.Loaded += MainWindow_Loaded;
            this.Activated += MainWindow_Activated;

            menuTabControl.SelectionChanged += MenuTabControl_SelectionChanged;
            codeTabControl.SelectionChanged += CodeTabControl_SelectionChanged;

            errorList.LineMoveRequest += ErrorList_LineMoveRequest;

            solutionExplorer.OpenRequest += SolutionExplorer_OpenRequest;
            solutionExplorer.OpenPropertyRequest += SolutionExplorer_OpenPropertyRequest;
            solutionExplorer.DeleteRequest += SolutionExplorer_DeleteRequest;

            BackBtn.Click += BackBtn_Click;

            Thread thr = new Thread(() =>
            {
                while (true)
                {
                    if (ParseSw.ElapsedMilliseconds > 500)
                    {
                        Stopwatch sw2 = new Stopwatch();
                        sw2.Start();

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
                            SetMessage("코드 파싱에 소요된 시간 : " + sw2.ElapsedMilliseconds.ToString() + "ms");
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

                                            var nFolding = list.First().Foldings.Select(i => new NewFolding(editor.Document.GetOffset(i.StartInt, 0),
                                                editor.Document.GetLineByNumber(i.EndInt).EndOffset)
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
                    }
                    Thread.Sleep(10);
                }
            });

            thr.SetApartmentState(ApartmentState.STA);
            thr.Start();


            var cbs = new CommandBinding[4] { null, null, null, null };

            RoutedCommand AddItem = new RoutedCommand();
            AddItem.InputGestures.Add(new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift));
            
            RoutedCommand AddMethod = new RoutedCommand();
            AddMethod.InputGestures.Add(new KeyGesture(Key.D, ModifierKeys.Control | ModifierKeys.Shift));

            RoutedCommand AddVar = new RoutedCommand();
            AddVar.InputGestures.Add(new KeyGesture(Key.V, ModifierKeys.Control | ModifierKeys.Shift));

            RoutedCommand SaveAll = new RoutedCommand();
            SaveAll.InputGestures.Add(new KeyGesture(Key.S, ModifierKeys.Control | ModifierKeys.Shift));
            

            cbs[0] = new CommandBinding(AddItem, Comm_ItmAdd);
            cbs[1] = new CommandBinding(AddMethod, Comm_MethodAdd);
            cbs[2] = new CommandBinding(AddVar, Comm_VarAdd);
            cbs[3] = new CommandBinding(SaveAll, Comm_SaveAll);

            this.CommandBindings.AddRange(cbs);
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
            
            if (connector != null)
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        foreach (VBComponentWrappingBase item in connector.GetFiles())
                        {
                            var itm = GetAllCodeTabs().Where(i => i.Header.ToString() == item.CompName).FirstOrDefault();
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

        #region [  윈도우 코드  ]
        Thread thr;
        public void SetMessage(string Message, int Delay = 3000)
        {
            thr?.Abort();
            thr = new Thread(() =>
            {
                Dispatcher.Invoke(() => { tbMessage.Text = Message; });

                Thread.Sleep(Delay);

                Dispatcher.Invoke(() => { tbMessage.Text = "준비"; });
            });

            thr.Start();
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

        private void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Opacity = 1.0;
            LoadComplete = true;

            this.Activate();
        }

        private void bg_DoWork(object sender, DoWorkEventArgs e)
        {
            MainDispatcher = Dispatcher;

            Dispatcher.Invoke(new Action(() =>
            {
                dbConnector = new SQLiteConnector();
                
                
                dbConnector.RecentFile_Get().ForEach((fl) => {
                    var itm = new RecentFileListViewItem(fl);
                    itm.OpenRequest += Itm_OpenRequest;
                    itm.CopyOpenRequest += Itm_CopyOpenRequest;
                    itm.DeleteRequest += Itm_DeleteRequest;
                    LVrecentFile.Items.Add(itm);
                });

                RunVersion.Text = VersionSelector.GetPPTVersion().GetDescription();
            }));
        }

        private void Itm_DeleteRequest(object sender)
        {
            if (!dbConnector.RecentFile_Remove(((RecentFileListViewItem)sender).FileLocation))
            {
                MessageBox.Show("알 수 없는 이유로 삭제에 실패했습니다.");
            }
            else
            {
                
                LVrecentFile.Items.Remove(sender);
            }
        }

        private void Itm_CopyOpenRequest(object sender)
        {
            RecentFileListViewItem itm = (RecentFileListViewItem)sender;
            SetMessage($"'{new FileInfo(itm.FileLocation).Name}' 프레젠테이션을 열고 있습니다.");
            InitalizeConnector(itm.FileLocation, true);
        }

        private void Itm_OpenRequest(object sender)
        {
            RecentFileListViewItem itm = (RecentFileListViewItem)sender;
            SetMessage($"'{new FileInfo(itm.FileLocation).Name}' 프레젠테이션을 열고 있습니다.");
            InitalizeConnector(itm.FileLocation);
        }

        public void AddCodeTab(string Name)
        {
            CodeEditor codeEditor = new CodeEditor();

            CloseableTabItem codeTab = new CloseableTabItem()
            {
                Header = Name,
                Content = codeEditor
            };

            codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
            codeEditor.SaveRequest += () => { SetMessage("저장되었습니다."); };
            codeTabControl.Items.Add(codeTab);
        }
        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            ParseSw.Restart();
            CodeSync(sender);
            btnUndo.IsEnabled = ((CodeEditor)sender).CanUndo;
            btnRedo.IsEnabled = ((CodeEditor)sender).CanRedo;
        }

        public void CodeSync(object sender)
        {
            int currLine = 0;
            if (sender?.GetType() == typeof(CodeEditor)) currLine = ((CodeEditor)sender).Text.SplitByNewLine().Count();

            projAnalyzer.CodeSync(connector.AllLineCount, connector.ComponentCount, currLine);
        }


        public CloseableTabItem FindCodeTab(string name)
        {
            return codeTabControl.Items.Cast<CloseableTabItem>().Where(i => i.Header.ToString() == name).FirstOrDefault();
        }

        public void AddCodeTab(VBComponentWrappingBase component)
        {
            var codeTab = codeTabControl.Items.Cast<CloseableTabItem>()
                             .Where(i => i.Header.ToString().ToLower() == component.CompName.ToLower())
                             .FirstOrDefault();

            if (codeTab != null)
            {
                codeTabControl.SelectedItem = codeTab;
                return;
            }

            var codeEditor = new CodeEditor(component)
            {
                Text = component.Code
            };
            
            codeTab = new CloseableTabItem()
            {
                Header = component.CompName,
                Content = codeEditor
            };

            codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
            codeEditor.TextChanged += CodeEditor_TextChanged;
            codeEditor.SaveRequest += () =>
            {
                SetMessage("저장되었습니다.");
                component.Code = codeEditor.Text;
                CodeSync(codeEditor);
            };

            codeEditor.RaiseFolding();

            codeTabControl.Items.Add(codeTab);
            codeTabControl.SelectedItem = codeTab;
            

            ParseSw.Restart();
        }

        public CodeEditor GetCodeTab()
        {
            if (codeTabControl.SelectedContent.GetType() == typeof(CodeEditor))
                return (CodeEditor)codeTabControl.SelectedContent;
            else
                return null;
        }

        public List<CodeEditor> GetAllCodeEditors()
        {
            List<CodeEditor> editorList = new List<CodeEditor>();

            foreach (CloseableTabItem tabItm in codeTabControl.Items)
                if (tabItm.Content.GetType() == typeof(CodeEditor))
                    editorList.Add((CodeEditor)tabItm.Content);

            return editorList;
        }

        public List<CloseableTabItem> GetAllCodeTabs()
        {
            List<CloseableTabItem> editorList = new List<CloseableTabItem>();

            foreach (CloseableTabItem tItm in codeTabControl.Items)
                if (tItm.Content.GetType() == typeof(CodeEditor))
                    editorList.Add(tItm);

            return editorList;
        }
        
        private string ProjName;

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

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            programTabControl.SelectedIndex = 0;
            this.NoTitle = false;
        }

        private void ProjectFileChange()
        {
            solutionExplorer.Update(connector);
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
        #endregion
        

        #region [  에디터 메뉴 탭  ]

        #region [  홈 탭 이벤트  ]


        #region [  클립보드  ]
        private void BtnCopy_SimpleButtonClicked(object sender)
        {
            Clipboard.Clear();
            Clipboard.SetText(((CodeEditor)codeTabControl.SelectedContent).SelectedText);
        }
        private void BtnPaste_SimpleButtonClicked(object sender)
        {
            if (Clipboard.ContainsText())
            {
                string t = Clipboard.GetText();
                CodeEditor editor = ((CodeEditor)codeTabControl.SelectedContent);

                if (editor.SelectionLength != 0) editor.SelectedText = t;
                else editor.TextArea.Document.Insert(editor.CaretOffset, t);
            }            
        }
        #endregion
        
        #region [  작업  ]
        private void BtnUndo_SimpleButtonClicked(object sender)
        {
            CodeEditor editor = ((CodeEditor)codeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanUndo) editor.Undo();
            btnUndo.IsEnabled = editor.CanUndo;
            btnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        private void BtnRedo_SimpleButtonClicked(object sender)
        {
            CodeEditor editor = ((CodeEditor)codeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanRedo) editor.Redo();
            btnUndo.IsEnabled = editor.CanUndo;
            btnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        #endregion
        
        #region [  슬라이드 관리  ]
        private void BtnNewSlide_SimpleButtonClicked(object sender)
        {
            int slideNumber = connector.Slide;

            connector.AddSlide(slideNumber + 1);


            switch (connector.Version)
            {
                case PPTVersion.PPT2010:
                    PPTConnector2010 conn2010 = (PPTConnector2010)connector;

                    if (conn2010.Presentation.Slides.Count != 0) slideNumber = conn2010.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2010.AddSlide(slideNumber + 1);
                    conn2010.Presentation.Slides.AddSlide(slideNumber + 1, conn2010.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2010.Presentation.Application.ActiveWindow.View.GotoSlide(slideNumber + 1);
                    break;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    PPTConnector2013 conn2013 = (PPTConnector2013)connector;

                    if (conn2013.Presentation.Slides.Count != 0) slideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2013.Presentation.Slides.AddSlide(slideNumber + 1, conn2013.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2013.Presentation.Application.ActiveWindow.View.GotoSlide(slideNumber + 1);
                    break;
                
            }

            SetMessage((slideNumber + 1) + "번째 슬라이드를 추가했습니다.");

        }

        private void BtnDelSlide_SimpleButtonClicked(object sender)
        {
            int SlideNumber = 0;
            if (connector.SlideCount == 0)
            {
                SetMessage("삭제할 슬라이드가 없습니다.");

                return;
            }

            if (MessageBox.Show(SlideNumber + "슬라이드를 삭제하시겠습니까?", "슬라이드 삭제 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                connector.DeleteSlide();
                SetMessage((SlideNumber) + "번째 슬라이드를 삭제했습니다.");
            }
        }

        #endregion


        private void BtnHelp_SimpleButtonClicked(object sender)
        {
            HelperWindow hWdw = new HelperWindow();
            hWdw.ShowDialog();
        }

        #endregion


        #region [  삽입 탭 이벤트  ]

        private void BtnAddClass_SimpleButtonClicked(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Class);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void BtnAddModule_SimpleButtonClicked(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Module);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void BtnAddForm_SimpleButtonClicked(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Form);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }


        private void BtnAddSub_SimpleButtonClicked(object sender)
        {
            var procWindow =
                new AddProcedureWindow(GetCodeTab(),
                                       ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                                       codeInfo, AddProcedureWindow.AddProcedureType.Sub);

            procWindow.ShowDialog();
        }

        private void BtnAddFunc_SimpleButtonClicked(object sender)
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(), 
                                   codeInfo, AddProcedureWindow.AddProcedureType.Function).ShowDialog();   
        }
        private void BtnAddProp_SimpleButtonClicked(object sender)
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(),
                                   codeInfo, AddProcedureWindow.AddProcedureType.Property).ShowDialog();
            
        }

        private void BtnAddMouseOverTrigger_SimpleButtonClicked(object sender)
        {
            new AddTriggerWindow(true,GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(connector);
        }

        private void BtnAddMouseClickTrigger_SimpleButtonClicked(object sender)
        {
            new AddTriggerWindow(false, GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(connector);
        }

        public string GetCodeTabName()
        {
            return ((TabItem)codeTabControl.SelectedItem).Header.ToString();
        }


        private void BtnAddVar_SimpleButtonClicked(object sender)
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, true).ShowDialog();
        }

        private void BtnAddConst_SimpleButtonClicked(object sender)
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, false).ShowDialog();
        }
        #endregion

        #region [  프로젝트 탭  ]
        
        private void PreDeclareFuncBtn_SimpleButtonClicked(object sender)
        {
            var tabItems = codeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "미리 정의된 함수").ToList();
            if (tabItems.Count >= 1)
            {
                codeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "미리 정의된 함수",
                    Content = new PreDeclareFuncManager() { Connector = connector }
                };
                codeTabControl.Items.Add(itm);
                codeTabControl.SelectedItem = itm;
            }
        }

        private void CheckError_SimpleButtonClicked(object sender)
        {
            var result = MessageBox.Show("코드 분석을 시작합니다.\r\n코드 분석은 현재 프로젝트에 있는 파일 모두를 분석해 오류를 확인합니다.\r\n" +
                            "저장되지 않은 내용은 검사되지 않으며 문법적 검사만 실행합니다.\r\n계속하시겠습니까?", "코드 분석 확인", MessageBoxButton.OKCancel);

            if (result == MessageBoxResult.OK)
            {
                var itm = connector.GetFiles();

                if (itm.Count != 0)
                {
                    ErrorWindow errWdw = new ErrorWindow(itm);

                    errWdw.ShowDialog();
                }
                else
                {
                    MessageBox.Show("파일이 없습니다!");
                }
            }
        }

        private void BtnFileSync_SimpleButtonClicked(object sender)
        {
            var itm = GetCodeTab();
            if (itm != null) itm.Save();
            SetMessage("저장되었습니다.");
        }

        private void BtnAllFileSync_SimpleButtonClicked(object sender)
        {
            var itm = GetAllCodeEditors();
            itm.ForEach(editor => editor.Save());
            SetMessage("전체 저장되었습니다.");
        }
        #endregion

        #endregion


        #region [  초기 화면  ]
        private void GridOpenAnotherPpt_MouseLeave(object sender, MouseEventArgs e)
        {
            TextBlock tb = null;
            if (((Control)sender).Name.ToLower().Contains("openanotherppt")) tb = TBOpenAnotherPpt;
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation")) tb = TBConnectPresentation;

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
            if (((Control)sender).Name.ToLower().Contains("openanotherppt")) tb = TBOpenAnotherPpt;
            else if (((Control)sender).Name.ToLower().Contains("connectpresentation")) tb = TBConnectPresentation;

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

                connector = tmpConn;

                CodeSync(null);
                programTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

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
                            tmpConn = new PPTConnector2010(FileLocation, CopyOpen);
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

                connector = tmpConn;
                
                CodeSync(null);
                programTabControl.SelectedIndex = 0;
                SetName();

            }), DispatcherPriority.Background);
        }

        private void SelectionChangedDetect()
        {
            projAnalyzer.CurrentShapeName = connector.SelectionShapeName;
        }

        private void ShapeChangedDetect()
        {
            projAnalyzer.ShapeCount = connector.Shapes().Count();
            projAnalyzer.CurrentShapeCount = connector.Shapes(connector.Slide).Count();
        }

        private void SlideChangedDetect()
        {
            projAnalyzer.SlideCount = connector.SlideCount;
        }

        #endregion

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
    }
}
