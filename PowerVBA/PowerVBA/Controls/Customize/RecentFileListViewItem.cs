using PowerVBA.Resources;
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

namespace PowerVBA.Controls.Customize
{
    [TemplatePart(Name = "TBFileName", Type = typeof(TextBlock))]
    [TemplatePart(Name = "TBFileLoc", Type = typeof(TextBlock))]

    class RecentFileListViewItem : ListViewItem
    {
        public delegate void BlankEventHandler();
        public event BlankEventHandler OpenRequest;
        public event BlankEventHandler CopyOpenRequest;
        public event BlankEventHandler CopyPathRequest;
        public event BlankEventHandler DeleteRequest;

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
        }

        #region [  이벤트 연결  ]



        private void MenuDelete_Click(object sender, RoutedEventArgs e)
        {
            DeleteRequest();
        }

        private void MenuCopyPath_Click(object sender, RoutedEventArgs e)
        {
            CopyPathRequest();
        }

        private void MenuCopyOpen_Click(object sender, RoutedEventArgs e)
        {
            CopyOpenRequest();
        }

        private void MenuOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenRequest();
        }

        #endregion


        TextBlock tbfn, tbfl;
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            tbfn = Template.FindName("TBFileName", this) as TextBlock;
            tbfl = Template.FindName("TBFileLoc", this) as TextBlock;

            var prop = DependencyPropertyDescriptor.FromProperty(TextBlock.TextProperty, typeof(TextBlock));
            prop.AddValueChanged(tbfn, FileNameChanged);
            prop.AddValueChanged(tbfl, FileNameChanged);

            FileNameChanged(this, null);
        }

        private void FileNameChanged(object sender, EventArgs e)
        {
            // 바탕화면 같은걸 실제 경로로 치환 해주기
            this.ToolTip = new ToolTip() { Content = Path.Combine(FolderLocation.Replace(" >> ","\\"), Content.ToString()) };
        }

        public static DependencyProperty FolderLocationProperty = DependencyHelper.Register();
        public static DependencyProperty SourceProperty = DependencyHelper.Register(new PropertyMetadata(ResourceImage.GetIconImage("EventIcon")));
        public string FolderLocation
        {
            get { return (string)GetValue(FolderLocationProperty); }
            set { SetValue(FolderLocationProperty, value); }
        }

        public ImageSource Source
        {
            get { return (ImageSource)GetValue(SourceProperty); }
            set { SetValue(SourceProperty, value); }
        }
    }
}
