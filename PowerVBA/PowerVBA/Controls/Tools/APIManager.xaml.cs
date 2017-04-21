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
    /// APIManager.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class APIManager : UserControl
    {
        public APIManager()
        {
            InitializeComponent();

            foreach (var itm in APIs.APIList)
            {
                APIListView.Items.Add(itm);
            }

            APIListView.SelectionChanged += APIListView_SelectionChanged;
        }

        private void APIListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var itm = ((APIInfo)APIListView.SelectedItem);

            RunFuncName.Text = itm.Name;
            TBDescription.Text = itm.APIStr + Environment.NewLine + itm.FullDescription;
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {

        }
    }

    public class test
    {
        public string Name { get; set; }
        public string ReturnData { get; set; }
        public string Description { get; set; }
        public APIInfo API { get; set; }
    }
}
