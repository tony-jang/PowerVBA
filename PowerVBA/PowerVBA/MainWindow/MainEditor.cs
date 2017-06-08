using PowerVBA.Controls.Customize;
using PowerVBA.Controls.Tools;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Core.Connector;
using PowerVBA.V2010.Connector;
using PowerVBA.V2013.Connector;
using PowerVBA.Windows;
using PowerVBA.Windows.AddWindows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace PowerVBA
{
    partial class MainWindow
    {
        #region [  홈 탭 이벤트  ]

        #region [  클립보드  ]
        private void BtnCopy_ButtonClick(object sender)
        {
            Clipboard.Clear();
            Clipboard.SetText(((CodeEditor)codeTabControl.SelectedContent).SelectedText);
        }
        private void BtnPaste_ButtonClick(object sender)
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
        private void BtnUndo_ButtonClick(object sender)
        {
            CodeEditor editor = ((CodeEditor)codeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanUndo) editor.Undo();
            btnUndo.IsEnabled = editor.CanUndo;
            btnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        private void BtnRedo_ButtonClick(object sender)
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
        private void BtnNewSlide_ButtonClick(object sender)
        {
            int slideNumber = connector.Slide;

            connector.AddSlide(slideNumber + 1);

            SetMessage((slideNumber + 1) + "번째 슬라이드를 추가했습니다.");

        }

        private void BtnDelSlide_ButtonClick(object sender)
        {
            int SlideNumber = 0;
            if (connector.SlideCount == 0)
            {
                SetMessage("삭제할 슬라이드가 없습니다.");

                return;
            }

            SlideNumber = connector.Slide;

            if (MessageBox.Show(SlideNumber + "슬라이드를 삭제하시겠습니까?", "슬라이드 삭제 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                connector.DeleteSlide();
                SetMessage((SlideNumber) + "번째 슬라이드를 삭제했습니다.");
            }
        }

        #endregion


        private void BtnHelp_ButtonClick(object sender)
        {
            HelperWindow hWdw = new HelperWindow();
            hWdw.ShowDialog();
        }

        #endregion


        #region [  삽입 탭 이벤트  ]

        private void BtnAddClass_ButtonClick(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Class);

            SolutionExplorer_Open(this, filewindow.ShowDialog());
        }

        private void BtnAddModule_ButtonClick(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Module);

            SolutionExplorer_Open(this, filewindow.ShowDialog());
        }

        private void BtnAddForm_ButtonClick(object sender)
        {
            AddFileWindow filewindow = new AddFileWindow(connector, AddFileWindow.AddFileType.Form);

            SolutionExplorer_Open(this, filewindow.ShowDialog());
        }


        private void BtnAddSub_ButtonClick(object sender)
        {
            var procWindow =
                new AddProcedureWindow(GetCodeTab(),
                                       ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                                       codeInfo, AddProcedureWindow.AddProcedureType.Sub);

            procWindow.ShowDialog();
        }

        private void BtnAddFunc_ButtonClick(object sender)
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(),
                                   codeInfo, AddProcedureWindow.AddProcedureType.Function).ShowDialog();
        }
        private void BtnAddProp_ButtonClick(object sender)
        {
            new AddProcedureWindow(GetCodeTab(), GetCodeTabName(),
                                   codeInfo, AddProcedureWindow.AddProcedureType.Property).ShowDialog();

        }

        private void BtnAddMouseOverTrigger_ButtonClick(object sender)
        {
            new AddTriggerWindow(true, GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(connector);
        }

        private void BtnAddMouseClickTrigger_ButtonClick(object sender)
        {
            new AddTriggerWindow(false, GetCodeTab(), codeInfo, GetCodeTabName()).ShowDialog(connector);
        }

        public string GetCodeTabName()
        {
            return ((TabItem)codeTabControl.SelectedItem).Header.ToString();
        }


        private void BtnAddVar_ButtonClick(object sender)
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, true).ShowDialog();
        }

        private void BtnAddConst_ButtonClick(object sender)
        {
            new AddVarWindow(GetCodeTab(), ((TabItem)codeTabControl.SelectedItem).Header.ToString(),
                             codeInfo, false).ShowDialog();
        }
        #endregion

        #region [  프로젝트 탭  ]

        private void PreDeclareFuncBtn_ButtonClick(object sender)
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

        private void CheckError_ButtonClick(object sender)
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

        private void BtnFileSync_ButtonClick(object sender)
        {
            var itm = GetCodeTab();
            if (itm != null) itm.Save();
            SetMessage("저장되었습니다.");
        }

        private void BtnAllFileSync_ButtonClick(object sender)
        {
            var itm = GetAllCodeEditors();
            itm.ForEach(editor => editor.Save());

            // 저장되지 않았을 경우
            if (!connector.IsLocalPresentation)
            {
                SetProgramTab(ProgramTabMenus.FileTab);
                SetFileTabMenu(FileTabMenus.SaveAs);
                return;
            }

            if (!(connector.Name.EndsWith(".ppsm") || connector.Name.EndsWith(".pptm")) && connector.ComponentCount != 0)
            {
                MessageBox.Show(".ppsm이거나 .pptm 형식이 아닌 파일은 매크로를 포함해 저장할 수 없습니다.\n" +
                    "다른 이름으로 저장하거나 매크로를 제거하세요.");
                return;
            }

            connector.Save();
            SetMessage("전체 저장되었습니다.");
        }

        /// <summary>
        /// 코드가 저장되었는지에 대한 여부를 가져옵니다. 저장되었을시 true를 하나라도 저장되지 않았을시 false를 반환합니다.
        /// </summary>
        /// <remarks>모든 코드 탭을 가져와서 Saved 속성을 통해 저장되지 않은 코드 탭의 갯수를 가져와 0개면 저장되었음을 반환</remarks>
        public bool CodeSaved => GetAllCodeTabs().Count(i => !i.Saved) == 0;
        #endregion
    }
}
