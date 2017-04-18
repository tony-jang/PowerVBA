using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using ICSharpCode.AvalonEdit.Rendering;
using PowerVBA.Core.AvalonEdit.CodeCompletion;
using PowerVBA.Core.AvalonEdit.Folding;
using PowerVBA.Core.AvalonEdit.Indentation;
using PowerVBA.Core.AvalonEdit.Parser;
using PowerVBA.Core.AvalonEdit.Renderer;
using PowerVBA.Core.AvalonEdit.Substitution;
using PowerVBA.Core.AvalonEdit.Substitution.Base;
using PowerVBA.Core.Error;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Xml;

using static PowerVBA.Global.Globals;
using static PowerVBA.Core.Extension.HighlightingRuleEx;
using PowerVBA.Core.Controls;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.AvalonEdit.Replace;

namespace PowerVBA.Core.AvalonEdit
{
    public class CodeEditor : TextEditor
    {
        #region [  전역 변수  ]

        protected List<BaseSubstitution> Substitutions = new List<BaseSubstitution>();

        protected CustomCompletionWindow compWindow;
        protected OverloadInsightWindow insightWindow;
        protected Thread thr;

        public FoldingManager foldingManager;
        protected VBAFoldingStrategy foldingStrategy;


        protected Stopwatch sw = new Stopwatch();
        protected CodeIndentation codeIndentation;
        protected List<CodeError> CodeErrors = new List<CodeError>();
        protected CustomToolTip errToolTip;
        protected double LineHeight;
        protected CodeParser codeParser;

        /// <summary>
        /// VBComponent와 연결되어 있는 여부를 가져옵니다.
        /// </summary>
        protected bool IsConnectedFile = false;
        protected VBComponentWrappingBase ConnectedFile = null;

        public static List<string> Classes;
        public static bool IsInitLibRead = false;

        public event BlankEventHandler SaveRequest;

        #endregion


        static CodeEditor()
        {
            Classes = new List<string>();

            if (!IsInitLibRead)
            {
                foreach (Assembly asm in new Assembly[] { Assembly.Load(Properties.Resources.LibPowerPoint),
                                                          Assembly.Load(Properties.Resources.Interop_VBA)})
                {
                    Type[] typelist = GetTypesInNamespace(asm);

                    foreach (var t in typelist)
                        if (t.IsInterface || t.IsEnum)
                        {
                            if (t.Name.StartsWith("_")) continue;
                            Classes.Add(t.Name);

                        }
                }
                

                IsInitLibRead = true;
            }
        }

        public void RaiseFolding()
        {
            sw.Restart();
        }

        public CodeEditor()
        {
            #region [  코드 폴딩  ]

            foldingManager = FoldingManager.Install(this.TextArea);
            foldingStrategy = new VBAFoldingStrategy();

            thr = new Thread(() => {
                do
                {
                    if (sw.ElapsedMilliseconds > 500)
                    {
                        sw.Reset();
                        //Dispatcher.Invoke(new Action(() => {
                        //    foldingStrategy.UpdateFoldings(foldingManager, this.Document);
                        //}), System.Windows.Threading.DispatcherPriority.Background);
                        
                    }
                    Thread.Sleep(10);
                } while (true);
            });

            thr.Start();
            
            this.TextChanged += delegate (object sender, EventArgs e)
            {
                sw.Restart();
            };

            #endregion

            #region [  하이라이팅, 키워드 연결  ]

            using (Stream s = new MemoryStream(Properties.Resources.VBA_Highlight))
            {
                using (XmlTextReader reader = new XmlTextReader(s))
                {
                    highlightingDefintion = HighlightingLoader.Load(reader, HighlightingManager.Instance);

                    this.SyntaxHighlighting = highlightingDefintion;
                }
            }

            foreach (string ClassName in Classes)
            {
                this.SyntaxHighlighting.MainRuleSet.Rules[1].Add(ClassName);
            }


            #endregion

            #region [  BackgroundRenderer 추가  ]

            this.TextArea.TextView.BackgroundRenderers.Add(new CurrentLineBackgroundRenderer(this));
            this.TextArea.TextView.BackgroundRenderers.Add(new ErrorLineBackgroundRenderer(this, CodeErrors));

            #endregion

            #region [  Substititions 초기화  ]

            Substitutions.Add(new VariableSubstitution(TextArea));
            Substitutions.Add(new MethodSubstitution(TextArea));
            Substitutions.Add(new ForSubstitution(TextArea));

            #endregion

            #region [  Indentation 초기화  ]

            codeIndentation = new CodeIndentation(this);

            #endregion

            #region [  변수 초기화  ]

            Completion = new VBACompletion();
            codeParser = new CodeParser(this, CodeErrors);

            #endregion
            
            #region [  옵션 변경  ]

            this.ShowLineNumbers = true;

            this.Options.InheritWordWrapIndentation = false;
            errToolTip = new CustomToolTip();

            this.FontSize = 13.333333333333333;
            this.FontFamily = new System.Windows.Media.FontFamily("DotumChe");


            LineHeight = this.TextArea.TextView.DefaultLineHeight;

            #endregion


            #region [  이벤트 연결  ]

            this.TextArea.TextEntering += OnTextEntering;
            this.TextArea.TextEntered += OnTextEntered;

            Thread ErrCheckthr = new Thread(CheckErrorPos);

            ErrCheckthr.Start();

            this.TextArea.Caret.PositionChanged += (sender, e) => this.TextArea.TextView.InvalidateLayer(KnownLayer.Background);

            this.MouseMove += CodeMouseMove;
            this.MouseLeave += CodeMouseLeave;
            this.PreviewKeyDown += PrevKeyDown;
            
            Application.Current.Exit += Current_Exit;

            #endregion

            #region [  커멘드 바인딩 연결  ]

            var ctrlSpace = new RoutedCommand();
            ctrlSpace.InputGestures.Add(new KeyGesture(Key.Space, ModifierKeys.Control));

            var ctrlS = new RoutedCommand();
            ctrlS.InputGestures.Add(new KeyGesture(Key.S, ModifierKeys.Control));

            var ctrlShiftZ = new RoutedCommand();
            ctrlShiftZ.InputGestures.Add(new KeyGesture(Key.Z, ModifierKeys.Control | ModifierKeys.Shift));

            var ctrlShiftS = new RoutedCommand();
            ctrlShiftS.InputGestures.Add(new KeyGesture(Key.S, ModifierKeys.Control | ModifierKeys.Shift));

            var ctrlF = new RoutedCommand();
            ctrlF.InputGestures.Add(new KeyGesture(Key.F, ModifierKeys.Control));

            var ctrlH = new RoutedCommand();
            ctrlH.InputGestures.Add(new KeyGesture(Key.H, ModifierKeys.Control));

            var cb1 = new CommandBinding(ctrlSpace, OnCtrlSpaceCommand);
            var cb2 = new CommandBinding(ctrlS, OnCtrlSCommand);
            var cb3 = new CommandBinding(ctrlShiftZ, OnCtrlShiftZCommand);
            var cb4 = new CommandBinding(ctrlShiftS, OnCtrlShiftSCommand);
            var cb5 = new CommandBinding(ctrlF, OnCtrlFCommand);
            var cb6 = new CommandBinding(ctrlH, OnCtrlHCommand);

            this.CommandBindings.Add(cb1);
            this.CommandBindings.Add(cb2);
            this.CommandBindings.Add(cb3);
            this.CommandBindings.Add(cb4);
            this.CommandBindings.Add(cb5);
            this.CommandBindings.Add(cb6);

            #endregion
        }

        private void OnCtrlHCommand(object sender, ExecutedRoutedEventArgs e)
        {
            FindReplaceDialog.ShowForReplace(this);
        }

        private void OnCtrlFCommand(object sender, ExecutedRoutedEventArgs e)
        {
            FindReplaceDialog.ShowForSearch(this);
        }

        public CodeEditor(VBComponentWrappingBase wrappingFile) : this() 
        {
            IsConnectedFile = true;
            ConnectedFile = wrappingFile;
        }


        private void OnCtrlShiftZCommand(object sender, ExecutedRoutedEventArgs e)
        {
            if (CanRedo) Redo();
        }

        private void OnCtrlShiftSCommand(object sender, ExecutedRoutedEventArgs e)
        {
            
        }

        private void OnCtrlSCommand(object sender, ExecutedRoutedEventArgs e)
        {
            RaiseSaveRequest();
        }

        public void RaiseSaveRequest()
        {
            Document.UndoStack.MarkAsOriginalFile();
            SaveRequest?.Invoke();
        }

        private static void Send(Key key, bool Shift = false, bool Ctrl = false)
        {
            if (Keyboard.PrimaryDevice != null)
            {
                if (Keyboard.PrimaryDevice.ActiveSource != null)
                {
                    if (Shift) key |= Key.LeftShift;
                    if (Ctrl) key |= Key.LeftCtrl;
                    var e = new KeyEventArgs(Keyboard.PrimaryDevice, Keyboard.PrimaryDevice.ActiveSource, 0, key)
                    {
                        RoutedEvent = Keyboard.KeyDownEvent
                    };
                    InputManager.Current.ProcessInput(e);

                    // Note: Based on your requirements you may also need to fire events for:
                    // RoutedEvent = Keyboard.PreviewKeyDownEvent
                    // RoutedEvent = Keyboard.KeyUpEvent
                    // RoutedEvent = Keyboard.PreviewKeyUpEvent
                }
            }
        }

        public void InputIndent()
        {
            this.Focus();
            Send(Key.Tab);
        }

        public void DeleteIndent()
        {
            this.Focus();
            Send(Key.Tab, true);
        }


        private void PrevKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                bool? b = compWindow?.IsVisible;
                bool booleanData;
                if (b.HasValue) booleanData = true;
                else booleanData = false;
                bool Handled = false;
                if (!booleanData)
                {
                    foreach (BaseSubstitution substitution in Substitutions)
                    {
                        substitution.Handled = false;
                        substitution.Convert();

                        if (substitution.Handled)
                        {
                            Handled = true;
                            break;
                        }
                    }
                }

                e.Handled = Handled;
            }
        }

        private void CompWindow_InsertionCompleted(object sender, ICompletionData data)
        {
            compWindow?.Close();
            compWindow = null;
            //bool Handled = false;

            foreach (BaseSubstitution substitution in Substitutions)
            {
                substitution.Handled = false;
                substitution.Convert();

                if (substitution.Handled)
                {
                    //Handled = true;
                    break;
                }
            }
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            thr.Abort();
        }

        int lastindex = 0;
        bool Handling;
        private void CodeMouseMove(object sender, MouseEventArgs e)
        {
            Handling = true;
        }


        public void CheckErrorPos()
        {
            do
            {
                if (!Handling) return;
                bool Handled = false;
                for (int ctr = 1; ctr <= this.LineCount; ctr++)
                {
                    double lineTop = this.TextArea.TextView.GetVisualTopByDocumentLine(ctr);
                    double MouseY = Mouse.GetPosition(this).Y;

                    if (MouseY >= lineTop && MouseY < LineHeight + lineTop)
                    {
                        string d = string.Join("\r\n", CodeErrors.Where((err) => err.LineNumber == ctr)
                                                                 .Select((err) => err.LineNumber + " : " + err.ErrorMessage));

                        int count = CodeErrors.Where((err) => err.LineNumber == ctr).Count();

                        // 현재 라인에 오류가 있을 경우
                        if (!string.IsNullOrEmpty(d))
                        {

                            if (lastindex != ctr)
                            {
                                errToolTip.IsOpen = false;
                                errToolTip.IsOpen = true;

                                errToolTip.Title = ctr + "번째 줄에서 발생한 " + count + "개의 오류";
                                errToolTip.Text = d;
                            }
                        }
                        else
                        {
                            errToolTip.IsOpen = false;

                        }
                        Handled = true;
                        lastindex = ctr;
                        break;
                    }
                }
                if (!Handled)
                {
                    {
                        errToolTip.IsOpen = false;
                        lastindex = -1;
                    }
                }
                Thread.Sleep(100);
            } while (true);
        }

        private void CodeMouseLeave(object sender, MouseEventArgs e)
        {
            Handling = false;
            errToolTip.IsOpen = false;
            lastindex = -1;
        }

        internal string GetIndentation(int Indentation)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < Indentation; i++)
            {
                sb.Append("\t");
            }

            return sb.ToString();
        }
        
        public VBACompletion Completion { get; set; }
        
        #region Code Completion
        private void OnTextEntered(object sender, TextCompositionEventArgs textCompositionEventArgs)
        {
            //if (this.CaretOffset >= 2)
            //{
            //    string prevStr = Text.Substring(this.CaretOffset - 2, 1);
            //    Console.WriteLine(prevStr);
            //    if (prevStr == " ") return;
            //}
            //ShowCompletion(textCompositionEventArgs.Text, false);

        }

        private void OnCtrlSpaceCommand(object sender, ExecutedRoutedEventArgs executedRoutedEventArgs)
        {
            ShowCompletion(null, true);
        }

        private void ShowCompletion(string enteredText, bool controlSpace)
        {

            //if (!controlSpace)
                // Console.WriteLine("Code Completion: TextEntered: " + enteredText);
            //else
                // Console.WriteLine("Code Completion: Ctrl+Space");
            
            if (Completion == null)
            {
                // Console.WriteLine("Code completion is null, cannot run code completion");
                return;
            }

            if (compWindow == null)
            {
                CodeCompletionResult results = null;
                try
                {
                    var doc = GetCompletionDocument(out int offset);
                    results = Completion.GetCompletions(doc, offset, controlSpace);
                }
                catch (Exception)
                {
                    // Console.WriteLine("Error in getting completion: " + exception);
                }
                if (results == null)
                    return;
                if (insightWindow == null && results.OverloadProvider != null)
                {
                    insightWindow = new OverloadInsightWindow(TextArea)
                    {
                        Provider = results.OverloadProvider
                    };
                    insightWindow.Show();
                    insightWindow.Closed += (o, args) => insightWindow = null;
                    return;
                }

                if (compWindow == null && results != null && results.CompletionData.Any())
                {
                    MiniLexer lexer = new MiniLexer(Text.Substring(0, CaretOffset));
                    lexer.Parse();

                    if (lexer.IsInString || lexer.IsInSingleComment || lexer.IsInVerbatimString) return;
                    if (enteredText == "\r" || enteredText == "\n") return;
                    // Open code completion after the user has pressed dot:
                    compWindow = new CustomCompletionWindow(TextArea)
                    {
                        CloseWhenCaretAtBeginning = controlSpace
                    };
                    this.compWindow.InsertionCompleted += CompWindow_InsertionCompleted;

                    compWindow.StartOffset -= results.TriggerWordLength;
                    //compWindow.EndOffset -= results.TriggerWordLength;

                    IList<ICompletionData> data = compWindow.CompletionList.CompletionData;
                    foreach (var completion in results.CompletionData.OrderBy(item => item.Text))
                    {
                        data.Add(completion);
                    }
                    if (results.TriggerWordLength > 0)
                    {
                        //compWindow.CompletionList.IsFiltering = false;
                        compWindow.CompletionList.SelectItem(results.TriggerWord);
                    }
                    compWindow.Show();
                    compWindow.Closed += (o, args) => compWindow = null;
                }
                if (compWindow != null)
                {
                    IList<ICompletionData> data = compWindow.CompletionList.CompletionData;
                    data.Clear();
                    foreach (var completion in results.CompletionData.OrderBy(item => item.Text))
                    {
                        data.Add(completion);
                    }
                    
                    if (compWindow.CompletionList.SelectedItem == null)
                    {
                        compWindow.CompletionList.SelectedItem = data.First();
                    }

                    //if (enteredText == " ") compWindow.Close();
                }
            }//end if

        }//end method

        private void OnTextEntering(object sender, TextCompositionEventArgs textCompositionEventArgs)
        {

            // Console.WriteLine("TextEntering: " + textCompositionEventArgs.Text);

            if (textCompositionEventArgs.Text == "\n")
                compWindow?.Close();
            if (textCompositionEventArgs.Text.Length > 0 && compWindow != null)
            {
                if (!char.IsLetterOrDigit(textCompositionEventArgs.Text[0]))
                {
                    // Whenever a non-letter is typed while the completion window is open,
                    // insert the currently selected element.
                    compWindow.CompletionList.RequestInsertion(textCompositionEventArgs);
                }
            }
            // Do not set e.Handled=true.
            // We still want to insert the character that was typed.
        }
        
        protected virtual IDocument GetCompletionDocument(out int offset)
        {
            offset = CaretOffset;
            return Document;
        }
        #endregion


    }


}