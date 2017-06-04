using Microsoft.Win32;
using PowerVBA.Codes.Extension;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Extension;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Resources;
using PowerVBA.Windows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerVBA
{
    // MainWindow Partial File :: Initalize Part
    // MainWindow 부분 파일 :: 초기화 부분
    // 동일한 프로그램에서 연결을 끊고 새로운 연결을 시작할때 사용됩니다.
    // 닫기나 다시 열기 등 모든 초기화 메소드들을 여기서 관리합니다.
    
    partial class MainWindow
    {
        /// <summary>
        /// 모든 내용을 초기화합니다. 모든 정보가 초기화됩니다.
        /// </summary>
        public void InitalizeAll()
        {
            // 커넥터 초기화
            connector = null;

            // 코드정보 클래스 초기화
            codeInfo = new Codes.CodeInfo();
            
            // 라이브러리 파일 목록 초기화
            libraryFiles = new List<FileInfo>();
            
            // 코드 탭 초기화
            codeTabControl.Items.Clear();

            ParseSw.Reset();
            
            // 솔루션 탐색기 초기화
            solutionExplorer.Reset();

            // 프로젝트 분석기 초기화
            projAnalyzer.Reset();

            // 오류 목록 초기화
            errorList.Reset();

            bg.RunWorkerAsync();
        }


        // 최근 파일 목록을 다시 작성합니다. 처음 화면과 파일 탭의 아이템에도 똑같이 적용됩니다.
        public void RecentFileSet()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                lvRecentFile.Items.Clear();
                lvRecentFile2.Items.Clear();

                dbConnector.RecentFileGet().ForEach((fl) => {
                    var itm = new RecentFileListViewItem(fl);
                    itm.OpenRequest += Itm_OpenRequest;
                    itm.CopyOpenRequest += Itm_CopyOpenRequest;
                    itm.DeleteRequest += Itm_DeleteRequest;
                    lvRecentFile.Items.Add(itm);

                    var itm2 = new ImageButton();

                    itm2.TextAlignment = TextAlignment.Left;
                    itm2.BackImage = ResourceImage.GetIconImage("PPTIcon");
                    itm2.ButtonMode = ImageButton.ButtonModes.LongWidth;
                    itm2.Content = new FileInfo(fl).Name;

                    lvRecentFile2.Items.Add(itm2);

                });

                RunVersion.Text = VersionSelector.GetPPTVersion().GetDescription();
            }));
        }


        public void RecentFolderSet()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                fileOpenRecentFolder.Items.Clear();


                dbConnector.RecentFolderGet().ForEach((fi) =>
                {
                    var itm = new ImageButton();

                    itm.TextAlignment = TextAlignment.Left;
                    itm.BackImage = ResourceImage.GetIconImage("PPtIcon");
                    itm.ButtonMode = ImageButton.ButtonModes.LongWidth;
                    itm.Content = fi.Replace("\\", " ≫ ");


                    fileOpenRecentFolder.Items.Add(itm);
                });

            }));
        }
    }
}
