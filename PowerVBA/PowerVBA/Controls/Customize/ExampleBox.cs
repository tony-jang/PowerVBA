using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    public class ExampleBox : Control
    {
        public static DependencyProperty TextProperty = DependencyHelper.Register();
        public string Text { get => (string)GetValue(TextProperty); set => SetValue(TextProperty, value); }
        

        public static DependencyProperty TitleProperty = DependencyHelper.Register();
        public string Title { get => (string)GetValue(TitleProperty); set => SetValue(TitleProperty, value); }

    }
}
