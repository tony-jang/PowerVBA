using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Folding;
using PowerVBA.Codes;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Extension;
using PowerVBA.Core.Wrap.WrapBase;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;


namespace PowerVBA
{
    // MainWindow Partial File :: Function Part
    // MainWindow 부분 파일 :: 함수 부분


    partial class MainWindow
    {
        #region [  Get or Find CodeTab(CodeEditor)  ]

        /// <summary>
        /// 코드 탭을 찾아 CloseableTabItem을 반환합니다.
        /// </summary>
        /// <param name="name">찾을 이름입니다.</param>
        /// <returns></returns>
        public CloseableTabItem FindCodeTab(string name)
        {
            return codeTabControl.Items.Cast<CloseableTabItem>().Where(i => i.Header.ToString() == name).FirstOrDefault();
        }

        /// <summary>
        /// 현재 선택되어 있는 코드 에디터를 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public CodeEditor GetCodeTab()
        {
            if (codeTabControl.SelectedContent.GetType() == typeof(CodeEditor))
                return (CodeEditor)codeTabControl.SelectedContent;
            else
                return null;
        }

        /// <summary>
        /// 모든 코드 에디터들을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<CodeEditor> GetAllCodeEditors()
        {
            return GetAllCodeTabs()
                .Select(i => i.Content as CodeEditor)
                .ToList();
        }

        public List<CodeEditor> GetNotSavedEditor()
        {
            return GetNotSavedTabs()
                .Select(i => i.Content as CodeEditor)
                .ToList();
        }

        /// <summary>
        /// 모든 코드 탭들을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<CloseableTabItem> GetAllCodeTabs()
        {
            List<CloseableTabItem> editorList = new List<CloseableTabItem>();

            foreach (CloseableTabItem tItm in codeTabControl.Items)
                if (tItm.Content.GetType() == typeof(CodeEditor))
                    editorList.Add(tItm);

            return editorList;
        }

        public List<CloseableTabItem> GetNotSavedTabs()
        {
            return GetAllCodeTabs()
                .Where(i => !i.Saved)
                .ToList();
        }

        #endregion

        #region [  Add CodeTab(CodeEditor)  ]

        /// <summary>
        /// VBComponentWrappingBase로부터 코드 탭을 추가합니다.
        /// </summary>
        /// <param name="component"></param>
        public void AddCodeTab(VBComponentWrappingBase component)
        {
            var codeTab = codeTabControl.Items.Cast<CloseableTabItem>()
                             .Where(i => i.Header.ToString().ToLower() == component.CompName.ToLower())
                             .FirstOrDefault();

            if (codeTab != null)
            {
                codeTabControl.SelectedItem = codeTab;
            }
            else
            {
                var codeEditor = new CodeEditor(component) { Text = component.Code };

                codeTab = new CloseableTabItem()
                {
                    Header = component.CompName,
                    Content = codeEditor
                };


                codeTab.SaveCloseRequest += CodeTab_SaveCloseRequest;

                codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Saved = (((UndoStack)sender).IsOriginalFile); };
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

                codeTabControl.Focus();
            }
            
            ParseSw.Restart();
        }

        private void CodeTab_SaveCloseRequest(object sender, EventArgs e)
        {
            if (sender is CloseableTabItem cItm)
            {
                if (cItm.Content is CodeEditor cEditor)
                {
                    cEditor.Save();
                }
            }
        }

        /// <summary>
        /// 코드 탭을 추가합니다.
        /// </summary>
        /// <param name="Name">추가할 이름을 입력합니다.</param>
        public void AddCodeTab(string Name)
        {
            CodeEditor codeEditor = new CodeEditor();

            CloseableTabItem codeTab = new CloseableTabItem()
            {
                Header = Name,
                Content = codeEditor
            };

            codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Saved = (((UndoStack)sender).IsOriginalFile); };
            codeEditor.SaveRequest += () => { SetMessage("저장되었습니다."); };
            codeTabControl.Items.Add(codeTab);
        }

        #endregion


        /// <summary>
        /// 코드를 동기화합니다.
        /// </summary>
        /// <param name="sender">CodeEdtor의 형태여야 합니다.</param>
        public void CodeSync(object sender)
        {
            int currLine = 0;
            if (sender is CodeEditor editor) currLine = editor.Text.SplitByNewLine().Count();
            

            

            projAnalyzer.CodeSync(connector.AllLineCount, connector.ComponentCount, currLine);
        }


        Thread thr;
        /// <summary>
        /// 윈도우 하단의 텍스트 블록에 메세지를 띄웁니다.
        /// </summary>
        /// <param name="Message">띄울 메세지 내용입니다.</param>
        /// <param name="Delay">띄울 시간입니다. 밀리초로 계산합니다.</param>
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


    }
}
