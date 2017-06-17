using Microsoft.Win32;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.Extension;
using PowerVBA.Core.Resources;
using PowerVBA.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ppt = Microsoft.Office.Interop.PowerPoint;

namespace PowerVBA
{
    partial class MainWindow
    {
        int OpenListViewLastIndex = -1;
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            var tc = (TabControl)sender;

            if (tc.SelectedIndex == OpenListViewLastIndex)
                return;

            OpenListViewLastIndex = tc.SelectedIndex;


            Console.WriteLine(tc.Name);

            if (tc.SelectedItem is TabItem ti)
            {
                switch (ti.Tag?.ToString())
                {
                    case "openAtPresentation":
                        RefreshPPTItem();
                        break;
                    case "openAtComputer":
                        RefreshOpenFolder();
                        break;
                    case "openAtRecent":
                        RecentFileSet();
                        break;
                }
            }
        }

        private void RefreshOpenFolder()
        {
            lvOpenFolders.Items.Clear();
            lvSaveAsFolder.Items.Clear();

            foreach (string folder in dbConnector.FolderTable.Get())
            {
                var openImgBtn = new ImageButton()
                {
                    TextAlignment = TextAlignment.Left,
                    ButtonMode = ImageButton.ButtonModes.LongWidth,
                    BackImage = ResourceImage.GetIconImage("FindFolderIcon")
                };

                var saveImgBtn = new ImageButton()
                {
                    TextAlignment = TextAlignment.Left,
                    ButtonMode = ImageButton.ButtonModes.LongWidth,
                    BackImage = ResourceImage.GetIconImage("FindFolderIcon")
                };

                saveImgBtn.Content = folder;
                saveImgBtn.ButtonClick += SaveImgBtn_ButtonClick;

                openImgBtn.Content = folder;
                openImgBtn.ButtonClick += FolderImgBtn_ButtonClick;

                lvOpenFolders.Items.Add(openImgBtn);
                lvSaveAsFolder.Items.Add(saveImgBtn);
            }
        }

        private void SaveImgBtn_ButtonClick(object sender)
        {
            SaveAsFile(((ImageButton)sender).Content.ToString());
        }

        // 폴더 
        private void FolderImgBtn_ButtonClick(object sender)
        {
            OpenAnotherPPT(((ImageButton)sender).Content.ToString());
        }
        // 프레젠테이션 연결
        private void ImgBtn_ButtonClick(object sender)
        {
            InitalizeAll();
            InitalizeConnector(new V2013.WrapClass.PresentationWrapping(((ppt.Presentation)((Control)sender).Tag)));
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

        private void BtnInfoSync_ButtonClick(object sender)
        {
            runFunc.Text = codeInfo.Functions.Count.ToString();
            runSub.Text = codeInfo.Subs.Count.ToString();
            runError.Text = codeInfo.ErrorList.Count.ToString();
        }

        public void OpenAnotherPPT(string initDirectory = "")
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "프레젠테이션|*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm"
            };
            if (!string.IsNullOrEmpty(initDirectory) && Directory.Exists(initDirectory))
            {
                ofd.InitialDirectory = initDirectory;
            }
            if (ofd.ShowDialog().Value)
            {
                tbProcessInfoTB.Text = "프레젠테이션을 열고 있습니다.";

                dbConnector.FileTable.Add(ofd.FileName);
                dbConnector.FolderTable.Add(new FileInfo(ofd.FileName).DirectoryName);

                InitalizeAll();
                InitalizeConnector(ofd.FileName);
            }
        }

        public void SaveAsFile(string initDirectory = "")
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "매크로 프레젠테이션 파일|*.pptm";

            if (!string.IsNullOrEmpty(initDirectory) && Directory.Exists(initDirectory))
            {
                sfd.InitialDirectory = initDirectory;
            }
            if (sfd.ShowDialog().Value)
            {
                var itm = GetAllCodeEditors();
                itm.ForEach(editor => editor.Save());

                if (connector.SaveAs(sfd.FileName))
                {
                    NameSet();
                    dbConnector.FileTable.Add(sfd.FileName);
                    dbConnector.FolderTable.Add(new FileInfo(sfd.FileName).DirectoryName);
                }
            }

            RefreshOpenFolder();
        }

        private void BtnSaveAsSetFile_ButtonClick(object sender)
        {
            SaveAsFile();
        }

        private void FileTabClose_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (CloseRequest())
            {
                AllClose();
            }

            e.Handled = true;
        }

        // 파일에서 뒤로 돌아가기 버튼 누를때
        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            SetProgramTab(ProgramTabMenu.MainEdit);
        }
    }
}
