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

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// SolutionExplorer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class SolutionExplorer : UserControl
    {
        public List<ListBox> ListBoxes = new List<ListBox>();
        public SolutionExplorer()
        {
            InitializeComponent();

            ListBoxes.Add(LBClass);
            ListBoxes.Add(LBModule);
            ListBoxes.Add(LBForms);
        }

        bool Handled = false;

        private void ListBoxes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Handled)
            {
                Handled = false;
                return;
            }

            foreach (ListBox lb in ListBoxes)
            {
                if (sender == lb) continue;

                if (lb.SelectedIndex != -1)
                {
                    Handled = true;
                    lb.SelectedIndex = -1;
                }
            }
        }
    }
}
