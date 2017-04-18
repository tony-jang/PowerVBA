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
using PowerVBA.V2013.WrapClass;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using System.Diagnostics;
using PowerVBA.Core.AvalonEdit.Parser;
using PowerVBA.Core.Error;
using ICSharpCode.AvalonEdit;
using exp = PowerVBA.Codes.Expressions;
using Microsoft.Vbe.Interop;
using PowerVBA.Codes.Expressions;
using PowerVBA.Core.Reference;
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
        PPTConnectorBase Connector;
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

            MenuTabControl.SelectionChanged += MenuTabControl_SelectionChanged;
            CodeTabControl.SelectionChanged += CodeTabControl_SelectionChanged;

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
                            CodeLists = CodeTabControl.Items
                                                      .Cast<TabItem>()
                                                      .Where(t => t.Content.GetType() == typeof(CodeEditor))
                                                      .Select(t => (((CodeEditor)t.Content).Text, t.Header.ToString())).ToList();
                        }));

                        new VBAParser(codeInfo).Parse(CodeLists);

                        errorList.Dispatcher.Invoke(new Action(() =>
                        {
                            this.Title = sw2.ElapsedMilliseconds.ToString();
                            errorList.SetError(codeInfo.ErrorList);
                        }));


                        Dispatcher.Invoke(new Action(() => {

                            foreach (CloseableTabItem itm in CodeTabControl.Items)
                            {
                                if (itm.Content.GetType() == typeof(CodeEditor))
                                {
                                    CodeEditor editor = itm.Content as CodeEditor;
                                    var list = codeInfo.Lines.Where(i => i.FileName == itm.Header.ToString());

                                    if (list.Count() >= 1)
                                    {

                                        var nFolding = list.First().Foldings.Select(i => new NewFolding(editor.Document.GetOffset(i.StartInt, 0),
                                            editor.Document.GetLineByNumber(i.EndInt).EndOffset) {
                                            Name = editor.Text.Substring(editor.Document.GetOffset(i.StartInt, 0),
                                                                         editor.Document.GetLineByNumber(i.StartInt).Length) });

                                        editor.foldingManager.UpdateFoldings(nFolding, -1);
                                    }

                                    
                                }
                            }
                        }));


                        //Dispatcher.Invoke(new Action(() => { this.Title = (((float)sw2.ElapsedTicks) / 1000).ToString(); }));
                        //Dispatcher.Invoke(new Action(() => { this.Title = codeInfo.ToString(); }));
                        //MessageBox.Show(codeInfo.ToString());
                        ParseSw.Reset();
                    }
                    Thread.Sleep(10);
                }
            });

            thr.SetApartmentState(ApartmentState.STA);
            thr.Start();

            RoutedCommand AddItem = new RoutedCommand();
            AddItem.InputGestures.Add(new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift));

            CommandBinding cb1 = new CommandBinding(AddItem, Comm_ItmAdd);

            this.CommandBindings.Add(cb1);
        }

        private void SolutionExplorer_DeleteRequest(object sender, VBComponentWrappingBase Data)
        {
            Connector.DeleteComponent(Data);
        }

        private void Comm_ItmAdd(object sender, ExecutedRoutedEventArgs e)
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddFileType.Module);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void ErrorList_LineMoveRequest(Codes.TypeSystem.Error err)
        {
            try
            {
                CloseableTabItem tabItm = CodeTabControl.Items.Cast<CloseableTabItem>()
                                                            .Where(i => i.Header.ToString() == err.FileName).FirstOrDefault();

                CodeEditor codeEditor = (CodeEditor)tabItm.Content;

                if (codeEditor == null || tabItm == null) return;

                codeEditor.ScrollToLine(err.Region.Begin.Line);
                codeEditor.SelectionLength = 0;
                codeEditor.CaretOffset = codeEditor.Document.GetOffset(err.Region.Begin);
                codeEditor.SelectionLength = codeEditor.Document.GetLineByOffset(codeEditor.SelectionStart).Length;

                CodeTabControl.SelectedItem = tabItm;
                codeEditor.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());   
            }
            
        }

        #region [  윈도우 코드  ]

        public void SetMessage(string Message, int Delay = 3000)
        {
            Thread thr = new Thread(() =>
            {
                Dispatcher.Invoke(() => { tbMessage.Text = Message; });

                Thread.Sleep(Delay);

                Dispatcher.Invoke(() => { tbMessage.Text = "준비"; });
            });

            thr.Start();
        }


        private void SolutionExplorer_OpenPropertyRequest()
        {
            var tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "솔루션 탐색기").ToList();
            if (tabItems.Count >= 1)
            {
                CodeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "솔루션 탐색기",
                    Content = new ProjectProperty(Connector)
                };
                CodeTabControl.Items.Add(itm);
                CodeTabControl.SelectedItem = itm;
            }
        }



        private void SolutionExplorer_OpenRequest(object sender, VBComponentWrappingBase Data)
        {
            if (Data != null)
            {
                var itm = Data.ToVBComponent2013();
                AddCodeTab(Data);
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
                    LVrecentFile.Items.Add(itm);
                });

                RunVersion.Text = VersionSelector.GetPPTVersion().GetDescription();
            }));

            //Environment.GetFolderPath(Environment.SpecialFolder.MyComputer),
            //Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            //Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            //Environment.GetFolderPath(Environment.SpecialFolder.System),
            //Environment.GetFolderPath(Environment.SpecialFolder.SystemX86),
            string[] Locations = { Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                                   };

            Stopwatch sw = new Stopwatch();
            sw.Start();

            List<string> CheckLocations = new List<string>();

            foreach (string s in Locations)
            {
                if (string.IsNullOrEmpty(s.Trim())) continue;
                CheckLocations.AddRange(DirectorySearcher.GetDirectories(s));
            }
            
            foreach(string s in CheckLocations)
            {
                try
                {
                    foreach (FileInfo f in new DirectoryInfo(s).GetFiles())
                    {
                        // dll tlb olb
                        switch (f.Extension)
                        {
                            case ".dll": case ".tlb": case ".olb":
                                LibraryFiles.Add(f);
                                break;
                        }
                    }
                }
                catch (UnauthorizedAccessException)
                { continue; }
            }
        }

        private void Itm_CopyOpenRequest()
        {
            
        }

        private void Itm_OpenRequest()
        {
            MessageBox.Show("!");
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
            CodeTabControl.Items.Add(codeTab);
        }
        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            ParseSw.Restart();
            CodeSync(sender);
            BtnUndo.IsEnabled = ((CodeEditor)sender).CanUndo;
            BtnRedo.IsEnabled = ((CodeEditor)sender).CanRedo;
        }

        public void CodeSync(object sender)
        {
            int lines = 0;
            Connector.ToConnector2013().VBProject.VBComponents.Cast<VBComponent>()
                                                              .ToList()
                                                              .ForEach((i) => lines += i.CodeModule.CountOfLines);

            int ComponentCount = Connector.ToConnector2013().VBProject.VBComponents.Count;

            int CurrLine = 0;
            if (sender?.GetType() == typeof(CodeEditor)) CurrLine = ((CodeEditor)sender).Text.SplitByNewLine().Count();

            projAnalyzer.CodeSync(lines, ComponentCount, CurrLine);
        }


        public void AddCodeTab(VBComponentWrappingBase component)
        {
            CodeEditor codeEditor = null;

            List<CloseableTabItem> tabItems;
            CloseableTabItem codeTab;

            switch (component.ClassVersion)
            {
                case PPTVersion.PPT2010:
                    var comp2010 = component.ToVBComponent2010();

                    tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString().ToLower() == comp2010.CodeModule.Name.ToLower() + GetExtensions(comp2010.Type)).ToList();

                    if (tabItems.Count >= 1)
                    {
                        CodeTabControl.SelectedItem = tabItems.First();
                        return;
                    }

                    codeEditor = new CodeEditor(component);

                    try
                    {
                        var module = comp2010.VBComponent.CodeModule;

                        if (comp2010.VBComponent.CodeModule.CountOfLines == 0) codeEditor.Text = "";
                        else codeEditor.Text = comp2010.VBComponent.CodeModule.get_Lines(1, comp2010.VBComponent.CodeModule.CountOfLines);
                    }
                    catch (Exception)
                    { MessageBox.Show("예외가 발생했습니다!"); }

                    codeTab = new CloseableTabItem()
                    {
                        Header = comp2010.Name + GetExtensions(comp2010.Type),
                        Content = codeEditor
                    };

                    CodeTabControl.Items.Add(codeTab);


                    codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
                    codeEditor.TextChanged += CodeEditor_TextChanged;
                    codeEditor.SaveRequest += () =>
                    {
                        SetMessage("저장되었습니다.");
                        comp2010.CodeModule.DeleteLines(1, comp2010.CodeModule.CountOfLines);
                        comp2010.CodeModule.AddFromString(codeEditor.Text);
                        CodeSync(codeEditor);
                    };

                    codeEditor.RaiseFolding();

                    CodeTabControl.SelectedItem = codeTab;

                    break;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    var comp2013 = component.ToVBComponent2013();

                    tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString().ToLower() == comp2013.CodeModule.Name.ToLower() + GetExtensions(comp2013.Type)).ToList();

                    if (tabItems.Count >= 1)
                    {
                        CodeTabControl.SelectedItem = tabItems.First();
                        return;
                    }

                    codeEditor = new CodeEditor(component);

                    try
                    {
                        var module = comp2013.VBComponent.CodeModule;

                        if (comp2013.VBComponent.CodeModule.CountOfLines == 0) codeEditor.Text = "";
                        else codeEditor.Text = comp2013.VBComponent.CodeModule.get_Lines(1, comp2013.VBComponent.CodeModule.CountOfLines);
                    }
                    catch (Exception)
                    { MessageBox.Show("예외가 발생했습니다!"); }

                    codeTab = new CloseableTabItem()
                    {
                        Header = comp2013.Name + GetExtensions(comp2013.Type),
                        Content = codeEditor
                    };

                    CodeTabControl.Items.Add(codeTab);


                    codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
                    codeEditor.TextChanged += CodeEditor_TextChanged;
                    codeEditor.SaveRequest += () =>
                    {
                        SetMessage("저장되었습니다.");
                        comp2013.CodeModule.DeleteLines(1, comp2013.CodeModule.CountOfLines);
                        comp2013.CodeModule.AddFromString(codeEditor.Text);
                        CodeSync(codeEditor);
                    };

                    codeEditor.RaiseFolding();

                    CodeTabControl.SelectedItem = codeTab;

                    break;
            }

            ParseSw.Restart();
        }

        public CodeEditor GetCodeTab()
        {
            if (CodeTabControl.SelectedContent.GetType() == typeof(CodeEditor))
                return (CodeEditor)CodeTabControl.SelectedContent;
            else
                return null;
        }

        public List<CodeEditor> GetAllCodeTabs()
        {
            List<CodeEditor> editorList = new List<CodeEditor>();
            foreach(CloseableTabItem tItm in CodeTabControl.Items)
            {
                if (tItm.Content.GetType() == typeof(CodeEditor)) editorList.Add((CodeEditor)tItm.Content);
            }

            return editorList;
        }

        public string GetExtensions(vbext_ComponentType Type)
        {
            switch (Type)
            {
                case vbext_ComponentType.vbext_ct_StdModule:
                    return ".bas";
                case vbext_ComponentType.vbext_ct_ClassModule:
                case vbext_ComponentType.vbext_ct_Document:
                    return ".cls";
                case vbext_ComponentType.vbext_ct_MSForm:
                    return ".frm";
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                default:
                    return "";
            }
        }

        private string ProjName;

        public void SetName(string customName = "")
        {
            if (customName != "")
            {
                ProjName = customName;
                this.Title = $"{customName} - PowerVBA";
            }
            else if (Connector != null)
            {
                PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                this.Title = conn2013.Presentation.Name + " - PowerVBA";

                ProjName = conn2013.Presentation.Name;

                if (conn2013.Presentation.ReadOnly == Microsoft.Office.Core.MsoTriState.msoTrue)
                    this.Title += " [읽기 전용]";
                
            }
        }

        private void PPTCloseDetect()
        {
            var result = MessageBox.Show(@"프레젠테이션이 PowerVBA의 코드가 저장되지 않은 상태에서 닫혔습니다.
코드 파일만 따로 저장하시겠습니까?
[예]를 누르시면 현재 열린 코드 파일만 추출되어 바탕화면에 저장됩니다.
[아니오]를 누르시면 코드 파일은 소멸되며 종료됩니다.", "PowerVBA", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    foreach (CloseableTabItem itm in CodeTabControl.Items)
                    {
                        if (itm.Content.GetType() == typeof(CodeEditor))
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
                }
                catch (Exception)
                { }
            }
            Environment.Exit(0);
        }

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            ProgramTabControl.SelectedIndex = 0;
            this.NoTitle = false;
        }

        private void ProjectFileChange()
        {
            solutionExplorer.Update(Connector);
        }


        private void CodeTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FindReplaceDialog.CloseWindow();

            if (CodeTabControl.SelectedItem == null)
            {
                ChangeBtnState(false, true);
                return;
            }

            CloseableTabItem tabItm = ((CloseableTabItem)CodeTabControl.SelectedItem);
            if (tabItm == null) return;
            if (tabItm.Content.GetType() == typeof(CodeEditor))
            {
                var editor = (CodeEditor)tabItm.Content;
                BtnUndo.IsEnabled = editor.CanUndo;
                BtnRedo.IsEnabled = editor.CanRedo;

                ChangeBtnState(true);
            }
            else
            {
                ChangeBtnState(false, true);
            }

            if (tabItm.Header.ToString().EndsWith(".bas") && tabItm.Content.GetType() == typeof(CodeEditor))
            {
                BtnAddMouseOverTrigger.IsEnabled = true;
                BtnAddMouseClickTrigger.IsEnabled = true;
            }
            else
            {
                BtnAddMouseOverTrigger.IsEnabled = false;
                BtnAddMouseClickTrigger.IsEnabled = false;
            }
            
        }

        public void ChangeBtnState(bool Flag, bool ContainUndoRedoButtons = false)
        {
            if (ContainUndoRedoButtons)
            {
                BtnUndo.IsEnabled = Flag;
                BtnRedo.IsEnabled = Flag;
            }
            BtnAddSub.IsEnabled = Flag;
            BtnAddFunc.IsEnabled = Flag;
            BtnAddProp.IsEnabled = Flag;

            BtnAddVar.IsEnabled = Flag;
            BtnAddConst.IsEnabled = Flag;

            BtnFileSync.IsEnabled = Flag;
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {

            bool Flag = false;
            if (Connector != null)
            {
                switch (VersionSelector.GetPPTVersion())
                {
                    case PPTVersion.PPT2010:
                        if (Connector.ToConnector2010().Presentation.Saved != Microsoft.Office.Core.MsoTriState.msoTrue) Flag = true;
                        break;
                    case PPTVersion.PPT2013:
                    case PPTVersion.PPT2016:
                        if (Connector.ToConnector2013().Presentation.Saved != Microsoft.Office.Core.MsoTriState.msoTrue) Flag = true;
                        break;
                    default:
                        break;
                }   
            }

            if (Flag)
            {
                MessageBoxResult result =
                               MessageBox.Show(Connector.Name + "에 저장되지 않은 내용이 있습니다. 저장하시겠습니까?",
                               "저장되지 않은 내용",
                               MessageBoxButton.YesNoCancel);

                switch (result)
                {
                    case MessageBoxResult.Yes:
                        if (!Connector.Save())
                        {
                            MessageBox.Show("저장에 실패했습니다. [읽기 전용]이거나 폰트가 없을 수 있습니다. 다른 이름으로 저장하기를 해주세요.", "저장 실패");
                            e.Cancel = true;
                            return;
                        }
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        e.Cancel = true;
                        return;
                }
            }

            AllClose();
        }

        public void AllClose()
        {
            Connector?.Dispose();
            Environment.Exit(0);
        }

        private void MenuTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MenuTabControl.SelectedIndex == 0)
            {
                ProgramTabControl.SelectedIndex = 2;
                MenuTabControl.SelectedIndex = 1;
                this.NoTitle = true;
            }
        }
        #endregion
        

        #region [  에디터 메뉴 탭  ]

        #region [  홈 탭 이벤트  ]


        #region [  클립보드  ]
        private void BtnCopy_SimpleButtonClicked()
        {
            Clipboard.Clear();
            Clipboard.SetText(((CodeEditor)CodeTabControl.SelectedContent).SelectedText);
        }
        private void BtnPaste_SimpleButtonClicked()
        {
            if (Clipboard.ContainsText())
            {
                string t = Clipboard.GetText();
                CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedContent);

                if (editor.SelectionLength != 0) editor.SelectedText = t;
                else editor.TextArea.Document.Insert(editor.CaretOffset, t);
            }            
        }
        #endregion
        
        #region [  작업  ]
        private void BtnUndo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanUndo) editor.Undo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        private void BtnRedo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanRedo) editor.Redo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        #endregion
        
        #region [  슬라이드 관리  ]
        private void BtnNewSlide_SimpleButtonClicked()
        {
            int SlideNumber = 0;


            switch (Connector.Version)
            {
                case PPTVersion.PPT2010:
                    PPTConnector2010 conn2010 = (PPTConnector2010)Connector;

                    if (conn2010.Presentation.Slides.Count != 0) SlideNumber = conn2010.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2010.Presentation.Slides.AddSlide(SlideNumber + 1, conn2010.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2010.Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
                    break;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                    if (conn2013.Presentation.Slides.Count != 0) SlideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2013.Presentation.Slides.AddSlide(SlideNumber + 1, conn2013.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2013.Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
                    break;
                
            }

        }

        private void BtnDelSlide_SimpleButtonClicked()
        {
            switch (VersionSelector.GetPPTVersion())
            {
                case PPTVersion.PPT2010:
                    PPTConnector2010 conn2010 = (PPTConnector2010)Connector;

                    int SlideNumber = conn2010.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                    if (MessageBox.Show(SlideNumber + "슬라이드를 삭제합니다. 계속하시려면 예로 계속하세요.", "슬라이드 삭제 확인",
                        MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        conn2010.Presentation.Slides[SlideNumber].Delete();
                    }
                    break;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                    int SlideNumber2 = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                    if (MessageBox.Show(SlideNumber2 + "슬라이드를 삭제합니다. 계속하시려면 예로 계속하세요.", "슬라이드 삭제 확인",
                        MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        conn2013.Presentation.Slides[SlideNumber2].Delete();
                    }
                    break;
            }
            
        }

        #endregion


        #endregion


        #region [  삽입 탭 이벤트  ]

        private void BtnAddClass_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddFileType.Class);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void BtnAddModule_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddFileType.Module);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }

        private void BtnAddForm_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddFileType.Form);

            SolutionExplorer_OpenRequest(this, filewindow.ShowDialog());
        }


        private void BtnAddSub_SimpleButtonClicked()
        {
            var procWindow =
                new AddProcedureWindow(GetCodeTab(),
                                       ((TabItem)CodeTabControl.SelectedItem).Header.ToString(),
                                       codeInfo, AddProcedureWindow.AddProcedureType.Sub);

            procWindow.ShowDialog();
        }

        private void BtnAddFunc_SimpleButtonClicked()
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(), 
                                   codeInfo, AddProcedureWindow.AddProcedureType.Function).ShowDialog();   
        }
        private void BtnAddProp_SimpleButtonClicked()
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(),
                                   codeInfo, AddProcedureWindow.AddProcedureType.Property).ShowDialog();
            
        }

        private void BtnAddMouseOverTrigger_SimpleButtonClicked()
        {
            new AddTriggerWindow(true,GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(Connector);
        }

        private void BtnAddMouseClickTrigger_SimpleButtonClicked()
        {
            new AddTriggerWindow(false, GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(Connector);
        }

        public string GetCodeTabName()
        {
            return ((TabItem)CodeTabControl.SelectedItem).Header.ToString();
        }


        private void BtnAddVar_SimpleButtonClicked()
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)CodeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, true).ShowDialog();
        }

        private void BtnAddConst_SimpleButtonClicked()
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)CodeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, false).ShowDialog();
        }
        #endregion

        #region [  프로젝트 탭  ]
        
        private void PreDeclareFuncBtn_SimpleButtonClicked()
        {
            var tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "미리 정의된 함수").ToList();
            if (tabItems.Count >= 1)
            {
                CodeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "미리 정의된 함수",
                    Content = new PreDeclareFuncManager() { Connector = Connector }
                };
                CodeTabControl.Items.Add(itm);
                CodeTabControl.SelectedItem = itm;
            }
        }

        private void CheckError_SimpleButtonClicked()
        {

        }

        private void BtnFileSync_SimpleButtonClicked()
        {
            var itm = GetCodeTab();
            if (itm != null) itm.RaiseSaveRequest();
            SetMessage("저장되었습니다.");
        }

        private void BtnAllFileSync_SimpleButtonClicked()
        {
            var itm = GetAllCodeTabs();
            itm.ForEach(editor => editor.RaiseSaveRequest());
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
                    catch (Exception)
                    {
                        MessageBox.Show("오류가 발생 했습니다.");
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
                PPTConnectorBase tmpConn;

                switch (VersionSelector.GetPPTVersion())
                {
                    case PPTVersion.PPT2010:
                        tmpConn = new PPTConnector2010(pptWrapping.ToPresentation2010());

                        tmpConn.PresentationClosed += PPTCloseDetect;
                        tmpConn.VBAComponentChange += ProjectFileChange;
                        tmpConn.SlideChanged += SlideChangedDetect;
                        tmpConn.ShapeChanged += ShapeChangedDetect;
                        tmpConn.SectionChanged += SectionChangedDetect;

                        Connector = tmpConn;
                        PropGrid.SelectedObject = Connector.ToConnector2010().Presentation;
                        break;
                    case PPTVersion.PPT2016:
                    case PPTVersion.PPT2013:
                        tmpConn = new PPTConnector2013(pptWrapping.ToPresentation2013());

                        tmpConn.PresentationClosed += PPTCloseDetect;
                        tmpConn.VBAComponentChange += ProjectFileChange;
                        tmpConn.SlideChanged += SlideChangedDetect;
                        tmpConn.ShapeChanged += ShapeChangedDetect;
                        tmpConn.SectionChanged += SectionChangedDetect;

                        Connector = tmpConn;
                        PropGrid.SelectedObject = Connector.ToConnector2013().Presentation;
                        break;
                }
                
                

                CodeSync(null);
                ProgramTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

        public void InitalizeConnector(string FileLocation = "")
        {
            Dispatcher.Invoke(new Action(() =>
            {
                this.NoTitle = false;
                PPTConnectorBase tmpConn;

                if (FileLocation != "" && !File.Exists(FileLocation))
                {
                    MessageBox.Show("파일 위치가 올바르지 않습니다.");
                    return;
                }
                switch (VersionSelector.GetPPTVersion())
                {
                    case PPTVersion.PPT2010:
                        if (FileLocation == string.Empty) tmpConn = new PPTConnector2010();
                        else tmpConn = new PPTConnector2010(FileLocation);

                        tmpConn.PresentationClosed += PPTCloseDetect;
                        tmpConn.VBAComponentChange += ProjectFileChange;
                        tmpConn.SlideChanged += SlideChangedDetect;
                        tmpConn.ShapeChanged += ShapeChangedDetect;
                        tmpConn.SectionChanged += SectionChangedDetect;

                        PropGrid.SelectedObject = tmpConn.ToConnector2010().Presentation;

                        break;
                    case PPTVersion.PPT2016:
                    case PPTVersion.PPT2013:
                        if (FileLocation == string.Empty) tmpConn = new PPTConnector2013();
                        else tmpConn = new PPTConnector2013(FileLocation);

                        tmpConn.PresentationClosed += PPTCloseDetect;
                        tmpConn.VBAComponentChange += ProjectFileChange;
                        tmpConn.SlideChanged += SlideChangedDetect;
                        tmpConn.ShapeChanged += ShapeChangedDetect;
                        tmpConn.SectionChanged += SectionChangedDetect;

                        PropGrid.SelectedObject = tmpConn.ToConnector2013().Presentation;

                        break;
                    default:
                        MessageBox.Show("지원하지 않는 버전입니다.");
                        return;
                }
                
                Connector = tmpConn;

                
                
                CodeSync(null);
                ProgramTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

        private void SectionChangedDetect()
        {
            //projAnalyzer.SectionCount = Connector.ToConnector2013().Presentation.SectionCount;
        }

        private void ShapeChangedDetect()
        {
            projAnalyzer.ShapeCount = Connector.Shapes().Count(); 
        }

        private void SlideChangedDetect()
        {
            projAnalyzer.SlideCount = Connector.SlideCount;
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

        private void DebugBtn_SimpleButtonClicked()
        {
            //CodeTabEditor
            
            var itm = ((CodeTabEditor)((CloseableTabItem)CodeTabControl.SelectedItem).Content).CodeEditor;
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
    }
}
