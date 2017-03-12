using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    class ImageTreeViewItem : TreeViewItem
    {
        public ImageTreeViewItem()
        {
            Style = FindResource("ImageTreeViewItemStyle") as Style;
        }

        public static DependencyProperty SourceProperty = DependencyHelper.Register();

        public ImageSource Source
        {
            get { return (ImageSource)GetValue(SourceProperty); }
            set { SetValue(SourceProperty, value); }
        }
    }
}
