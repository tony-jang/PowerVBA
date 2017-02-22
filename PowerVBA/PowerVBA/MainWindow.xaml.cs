using ICSharpCode.AvalonEdit.Folding;
using PowerVBA.Core.AvalonEdit.Folding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Threading;
using PowerVBA.Core.AvalonEdit.Indentation;
using System.IO;
using System.Reflection;
using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Rendering;
using PowerVBA.Core.AvalonEdit.Renderer;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Core.Extension;
using System.Text.RegularExpressions;
using static PowerVBA.Core.AvalonEdit.Convert.StringConverter;

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        FoldingManager foldingManager;
        VBAFoldingStrategy foldingStrategy;

        Thread thr;

        Stopwatch sw = new Stopwatch();

        public MainWindow()
        {
            InitializeComponent();
            foldingManager = FoldingManager.Install(codeEditor.TextArea);
            foldingStrategy = new VBAFoldingStrategy();


            #region [  코드 폴딩  ]
            thr = new Thread(() => {
                do
                {
                    if (sw.ElapsedMilliseconds > 500)
                    {
                        sw.Reset();
                        Dispatcher.Invoke(new Action(() => { foldingStrategy.UpdateFoldings(foldingManager, codeEditor.Document);
                        }), System.Windows.Threading.DispatcherPriority.Background);   
                    }
                    Thread.Sleep(10);
                } while (true);
            });

            codeEditor.TextChanged += delegate (object sender, EventArgs e)
            {
                if (!(thr.ThreadState.HasFlag(System.Threading.ThreadState.Running) || thr.ThreadState.HasFlag(System.Threading.ThreadState.WaitSleepJoin))) thr.Start();
                sw.Restart();
            };

            #endregion


            #region [  Indent  ]

            codeEditor.TextArea.IndentationStrategy = new VBAIndentationStrategy();

            #endregion


            #region [  하이라이팅 연결  ]

            using (Stream s = new MemoryStream(PowerVBA.Properties.Resources.VBA_Highlight))
            {
                using (XmlTextReader reader = new XmlTextReader(s))
                {
                    codeEditor.SyntaxHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }

            #endregion


            #region [  BackgroundRenderer 추가  ]

            codeEditor.TextArea.TextView.BackgroundRenderers.Add(new CurrentLineBackgroundRenderer(codeEditor));

            #endregion


            #region [  이벤트 핸들러 연결  ]

            this.Closing += WindowClosingEvent;


            // codeEditor
            codeEditor.TextArea.Caret.PositionChanged += (sender, e) => codeEditor.TextArea.TextView.InvalidateLayer(KnownLayer.Background);
            codeEditor.TextChanged += codeChanged;
            codeEditor.PreviewKeyDown += prevKeyDown;

            #endregion


            #region [  TextEditor 옵션 변경  ]


            codeEditor.Options.InheritWordWrapIndentation = false;
            //codeEditor.Options.HighlightCurrentLine = true;

            #endregion  

        }




        #region [  Avalon Text Editor 이벤트  ]
        private void codeChanged(object sender, EventArgs e)
        {
            
        }

        string lineStartPattern = @"^(public|private) (sub|type|function) ([_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)$";

        private void prevKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                DocumentLine line = codeEditor.Document.GetLineByOffset(codeEditor.CaretOffset);

                string clonestr = codeEditor.Text.Clone().ToString();

                string codeText = (clonestr.Substring(line.Offset, line.Length));

                int Offset = line.Offset, Length = line.Length;

                if (Regex.IsMatch(codeText,lineStartPattern))
                {
                    string tmp;
                    Match m = Regex.Match(codeText, lineStartPattern);

                    string Accessor = m.Groups[1].Value;
                    string Type = m.Groups[2].Value;
                    string Name = m.Groups[3].Value;

                    tmp = codeEditor.Text.Insert(line.Offset + line.Length, "\r\n\t\r\nEnd " + Type);

                    tmp = tmp.Change(line.Offset, line.Length, $"{ConvertAccessor(Accessor)} {ConvertType(Type)} {Name}");

                    codeEditor.Text = tmp;

                    codeEditor.TextArea.Caret.Offset = Offset + Length + 3;

                    e.Handled = true;
                }

                
            }
        }
        #endregion


        #region [  윈도우 이벤트  ]
        private void WindowClosingEvent(object sender, System.ComponentModel.CancelEventArgs e)
        {
            thr.Abort();
        }
        #endregion

    }
}
