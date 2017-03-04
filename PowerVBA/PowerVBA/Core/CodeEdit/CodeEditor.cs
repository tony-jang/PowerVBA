using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using ICSharpCode.AvalonEdit.Rendering;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.CodeEdit.CodeCompletion;
using PowerVBA.Core.CodeEdit.Folding;
using PowerVBA.Core.CodeEdit.Indentation;
using PowerVBA.Core.CodeEdit.Parser;
using PowerVBA.Core.CodeEdit.Renderer;
using PowerVBA.Core.CodeEdit.Substitution;
using PowerVBA.Core.CodeEdit.Substitution.Base;
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

namespace PowerVBA.Core.CodeEdit
{
    class CodeEditor : TextEditor
    {
        #region [  전역 변수  ]

        protected List<BaseSubstitution> Substitutions = new List<BaseSubstitution>();

        protected CustomCompletionWindow compWindow;
        protected OverloadInsightWindow insightWindow;
        protected Thread thr;

        protected FoldingManager foldingManager;
        protected VBAFoldingStrategy foldingStrategy;


        protected Stopwatch sw = new Stopwatch();
        protected CodeIndentation codeIndentation;
        protected List<CodeError> CodeErrors = new List<CodeError>();
        protected ErrorToolTip errToolTip;
        protected double LineHeight;

        public static List<string> Classes;
        public static bool IsInitLibRead = false;

        #endregion
        static CodeEditor()
        {
            Classes = new List<string>();

            if (!IsInitLibRead)
            {
                Type[] typelist = GetTypesInNamespace(Assembly.Load(PowerVBA.Properties.Resources.LibPowerPoint), "Microsoft.Office.Interop.PowerPoint");

                foreach (var t in typelist)
                    if (t.IsInterface || t.IsEnum)
                    {
                        if (t.Name.StartsWith("_")) continue;
                        Classes.Add(t.Name);
                        
                    }
                IsInitLibRead = true;
            }
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
                        Dispatcher.Invoke(new Action(() => {
                            foldingStrategy.UpdateFoldings(foldingManager, this.Document);
                        }), System.Windows.Threading.DispatcherPriority.Background);
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

            using (Stream s = new MemoryStream(PowerVBA.Properties.Resources.VBA_Highlight))
            {
                using (XmlTextReader reader = new XmlTextReader(s))
                {
                    highlightingDefintion = HighlightingLoader.Load(reader, HighlightingManager.Instance);

                    this.SyntaxHighlighting = highlightingDefintion;
                }
            }

            foreach(string ClassName in Classes)
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

            #endregion

            #region [  Indentation 초기화  ]

            codeIndentation = new CodeIndentation(this);

            #endregion

            #region [  변수 초기화  ]

            Completion = new VBACompletion();
            CodeParser codeParser = new CodeParser(this, CodeErrors);

            #endregion
            
            #region [  옵션 변경  ]

            this.ShowLineNumbers = true;

            this.Options.InheritWordWrapIndentation = false;
            errToolTip = new ErrorToolTip();

            this.FontSize = 13.333333333333333;

            LineHeight = this.TextArea.TextView.DefaultLineHeight;

            #endregion

            #region [  이벤트 연결  ]

            this.TextArea.TextEntering += OnTextEntering;
            this.TextArea.TextEntered += OnTextEntered;
            this.TextArea.SelectionChanged += codeSelChanged;

            this.TextArea.Caret.PositionChanged += (sender, e) => this.TextArea.TextView.InvalidateLayer(KnownLayer.Background);
            this.TextArea.Caret.PositionChanged += codeCaretChanged;
            
            this.MouseMove += codeMouseMove;
            this.MouseLeave += codeMouseLeave;
            this.PreviewKeyDown += PrevKeyDown;
            
            Application.Current.Exit += Current_Exit;

            #endregion

            #region [  커멘드 바인딩 연결  ]

            var ctrlSpace = new RoutedCommand();
            ctrlSpace.InputGestures.Add(new KeyGesture(Key.Space, ModifierKeys.Control));
            var cb = new CommandBinding(ctrlSpace, OnCtrlSpaceCommand);

            this.CommandBindings.Add(cb);

            #endregion
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

        private void compWindow_InsertionCompleted(object sender, ICompletionData data)
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
        private void codeMouseMove(object sender, MouseEventArgs e)
        {
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
        }


        private void codeMouseLeave(object sender, MouseEventArgs e)
        {
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



        private void codeSelChanged(object sender, EventArgs e)
        {
            // 텍스트 선택이 바뀌었을때 발생 (단순히 커서를 움직이는 것 만으로는 발생하지 않음)
        }

        DocumentLine LastLine = null;
        private void codeCaretChanged(object sender, EventArgs e)
        {
            DocumentLine line = this.Document.GetLineByOffset(this.CaretOffset);

            if (LastLine != line)
            {
                //codeIndentation.Indent();
                // 선택된 라인이 바뀌었을때 (예 : 선택된 라인이 4였다가 5나 7로 바뀌었을때)
                LastLine = line;
            }
        }

        
        public VBACompletion Completion { get; set; }
        
        #region Code Completion
        private void OnTextEntered(object sender, TextCompositionEventArgs textCompositionEventArgs)
        {
            ShowCompletion(textCompositionEventArgs.Text, false);

        }

        private void OnCtrlSpaceCommand(object sender, ExecutedRoutedEventArgs executedRoutedEventArgs)
        {
            ShowCompletion(null, true);
        }

        private void ShowCompletion(string enteredText, bool controlSpace)
        {
            if (!controlSpace)
                Console.WriteLine("Code Completion: TextEntered: " + enteredText);
            else
                Console.WriteLine("Code Completion: Ctrl+Space");
            
            if (Completion == null)
            {
                Console.WriteLine("Code completion is null, cannot run code completion");
                return;
            }

            if (compWindow == null)
            {
                CodeCompletionResult results = null;
                try
                {
                    var offset = 0;
                    var doc = GetCompletionDocument(out offset);
                    results = Completion.GetCompletions(doc, offset, controlSpace);
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Error in getting completion: " + exception);
                }
                if (results == null)
                    return;
                if (insightWindow == null && results.OverloadProvider != null)
                {
                    insightWindow = new OverloadInsightWindow(TextArea);
                    insightWindow.Provider = results.OverloadProvider;
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
                    compWindow = new CustomCompletionWindow(TextArea);
                    compWindow.CloseWhenCaretAtBeginning = controlSpace;

                    this.compWindow.InsertionCompleted += compWindow_InsertionCompleted;

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
                }
            }//end if

        }//end method

        private void OnTextEntering(object sender, TextCompositionEventArgs textCompositionEventArgs)
        {

            Console.WriteLine("TextEntering: " + textCompositionEventArgs.Text);

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

        /// <summary>
        /// Gets the document used for code completion, can be overridden to provide a custom document
        /// </summary>
        /// <param name="offset"></param>
        /// <returns>The document of this text editor.</returns>
        protected virtual IDocument GetCompletionDocument(out int offset)
        {
            offset = CaretOffset;
            return Document;
        }
        #endregion


    }


}