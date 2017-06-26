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
    public class ImageTreeViewItem : TreeViewItem
    {
        public ImageTreeViewItem()
        {
            Style = FindResource("ImageTreeViewItemStyle") as Style;
        }

        public static DependencyProperty SourceProperty = DependencyHelper.Register();
        public static DependencyProperty ImageHeightProperty = DependencyHelper.Register();


        public ImageSource Source
        {
            get => (ImageSource)GetValue(SourceProperty);
            set => SetValue(SourceProperty, value);
        }
        public double ImageHeight
        {
            get => (double)GetValue(ImageHeightProperty);
            set => SetValue(SourceProperty, value);
        }
    }
}
