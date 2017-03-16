using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WPFExtension;

namespace PowerVBA.Core.Controls
{
    /// <summary>
    /// CodeBlock.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class CodeBlock : UserControl
    {
        public CodeBlock()
        {
            InitializeComponent();

            this.MouseLeftButtonDown += CodeBlock_MouseLeftButtonDown;

            var selProp = DependencyPropertyDescriptor.FromProperty(CodeBlock.IsSelectedProperty, typeof(CodeBlock));
            var selProp2 = DependencyPropertyDescriptor.FromProperty(CodeBlock.IsFocusedProperty, typeof(CodeBlock));
            selProp.AddValueChanged(this, ChangeHandler);
            selProp2.AddValueChanged(this, FocusHandler);
        }

        private void CodeBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.Focus();
        }

        private void FocusHandler(object sender, EventArgs e)
        {
            if (this.IsFocused) this.IsSelected = true;
            if (!this.IsFocused) this.IsSelected = false;
        }

        private void ChangeHandler(object sender, EventArgs e)
        {
            if (this.IsSelected)
            {
                this.Background = (new BrushConverter()).ConvertFromString("#FFD24726") as Brush;
                this.Foreground = Brushes.White;
            }
            else
            {
                this.Background = Brushes.White;
                this.Foreground = Brushes.Black;
            }
        }

        public static DependencyProperty BlockTypeProperty = DependencyHelper.Register();

        public CodeBlockType BlockType
        {
            get { return (CodeBlockType)GetValue(BlockTypeProperty); }
            set { SetValue(BlockTypeProperty, value); }
        }
        public static DependencyProperty IsSelectedProperty = DependencyHelper.Register();

        public bool IsSelected
        {
            get { return (bool)GetValue(IsSelectedProperty); }
            set { SetValue(IsSelectedProperty, value); }
        }
    }

    public enum CodeBlockType
    {
        Method,
        ReturnMethod,
        Property
    }

}
