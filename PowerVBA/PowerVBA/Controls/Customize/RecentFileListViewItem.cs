﻿using PowerVBA.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WPFExtension;
using static PowerVBA.Global.Globals;

namespace PowerVBA.Controls.Customize
{
    [TemplatePart(Name = "TBFileName", Type = typeof(TextBlock))]
    [TemplatePart(Name = "TBFileLoc", Type = typeof(TextBlock))]
    [TemplatePart(Name = "mainGrid", Type = typeof(Grid))]

    class RecentFileListViewItem : ListViewItem
    {
        public delegate void BlankEventHandler();
        public event SenderEventHandler OpenRequest;
        public event SenderEventHandler CopyOpenRequest;
        public event SenderEventHandler DeleteRequest;

        public RecentFileListViewItem()
        {
            this.Style = FindResource("RecentFileListViewItemStyle") as Style;

            MenuItem menuOpen = new MenuItem() { Header = "열기" };
            MenuItem menuCopyOpen = new MenuItem() { Header = "복사본 열기" };
            MenuItem menuCopyPath = new MenuItem() { Header = "클립보드에 경로 복사" };
            MenuItem menuDelete = new MenuItem() { Header = "목록에서 제거" };

            ContextMenu = new ContextMenu();

            menuOpen.Click += MenuOpen_Click;
            menuCopyOpen.Click += MenuCopyOpen_Click;
            menuCopyPath.Click += MenuCopyPath_Click;
            menuDelete.Click += MenuDelete_Click;

            ContextMenu.Items.Add(menuOpen);
            ContextMenu.Items.Add(menuCopyOpen);
            ContextMenu.Items.Add(menuCopyPath);
            ContextMenu.Items.Add(menuDelete);

            
            MouseDown += RecentFileListViewItem_MouseDown;
        }

        private void RecentFileListViewItem_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                OpenRequest?.Invoke(this);
            }
        }

        public RecentFileListViewItem(string FileName) : this()
        {
            FileLocation = FileName;
        }

        #region [  이벤트 연결  ]



        private void MenuDelete_Click(object sender, RoutedEventArgs e)
        {
            DeleteRequest?.Invoke(this);
        }

        private void MenuCopyPath_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(FileLocation);
        }

        private void MenuCopyOpen_Click(object sender, RoutedEventArgs e)
        {
            CopyOpenRequest?.Invoke(this);
        }

        private void MenuOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenRequest?.Invoke(this);
        }

        #endregion


        TextBlock TBFileName, TBFileLocation;
        Grid mainGrid;
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            TBFileName = Template.FindName("TBFileName", this) as TextBlock;
            TBFileLocation = Template.FindName("TBFileLoc", this) as TextBlock;
            mainGrid = Template.FindName("mainGrid", this) as Grid;
            mainGrid.MouseDown += RecentFileListViewItem_MouseDown;

            var prop = DependencyPropertyDescriptor.FromProperty(FileLocationProperty, typeof(RecentFileListViewItem));
            prop.AddValueChanged(TBFileName, FileNameChanged);
            prop.AddValueChanged(TBFileLocation, FileNameChanged);

            FileNameChanged(this, null);
        }

        private void FileNameChanged(object sender, EventArgs e)
        {
            // 바탕화면 같은걸 실제 경로로 치환 해주기
            // this.ToolTip = new ToolTip() { Content = Path.Combine(FileLocation.Replace(" >> ","\\"), Content.ToString()) };
            if (File.Exists(FileLocation))
            {
                FileInfo fi = new FileInfo(FileLocation);

                string Dir = fi.DirectoryName;
                string File = fi.Name;

                TBFileLocation.Text = Dir.Replace("\\", " ≫ ");

                TBFileName.Text = File;

                this.ToolTip = new ToolTip() { Content = FileLocation };
            }
            else
            {
                TBFileName.Text = "존재하지 않는 파일입니다.";
                TBFileLocation.Text = "존재하지 않는 파일입니다.";

                this.ToolTip = null;
            }
        }

        public static DependencyProperty FileLocationProperty = DependencyHelper.Register();
        public static DependencyProperty SourceProperty = DependencyHelper.Register(new PropertyMetadata(ResourceImage.GetIconImage("EventIcon")));
        

        public ImageSource Source
        {
            get { return (ImageSource)GetValue(SourceProperty); }
            set { SetValue(SourceProperty, value); }
        }

        public string FileLocation
        {
            get { return (string)GetValue(FileLocationProperty); }
            set { SetValue(FileLocationProperty, value); }
        }

        public string FileName => TBFileName?.Text;
    }
}
