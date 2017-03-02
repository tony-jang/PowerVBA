using ICSharpCode.AvalonEdit.Folding;
using PowerVBA.Core.CodeEdit.Folding;
using System;
using System.Collections.Generic;
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
using System.Diagnostics;
using System.Threading;
using PowerVBA.Windows;
using ICSharpCode.AvalonEdit;
using PowerVBA.Core.CodeEdit;

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : ChromeWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            


            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;

            DebuggingAdd();
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MainTabControl.SelectedIndex == 0) MainTabControl.SelectedIndex = 1;
        }


        #region [  디버그용 코드  ]

        public void DebuggingAdd()
        {
            int index = 0;
            foreach(string d in CodeEditor.Classes)
            {
                TextBlock tb = new TextBlock { Text = d , ToolTip = d };
                HighlightingClassesListBox.Items.Add(tb);
                tb.Width = 110;
                index++;
            }
            
        }

        #endregion
    }
}
