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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    [TemplatePart(Name = "TBFileName", Type = typeof(TextBlock))]
    [TemplatePart(Name = "TBFileLoc", Type = typeof(TextBlock))]

    class RecentFileListViewItem : ListViewItem
    {
        
        public RecentFileListViewItem()
        {
            this.Style = FindResource("RecentFileListViewItemStyle") as Style;
            
        }
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
