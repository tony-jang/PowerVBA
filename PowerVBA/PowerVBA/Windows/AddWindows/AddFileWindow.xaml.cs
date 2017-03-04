using PowerVBA.Core.Connector;
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
using System.Windows.Shapes;

namespace PowerVBA.Windows.AddWindows
{
    /// <summary>
    /// AddFileWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AddFileWindow : ChromeWindow
    {
        public AddFileWindow(PPTConnector connector, bool IsClass)
        {
            InitializeComponent();
            if (IsClass) btnClass.IsChecked = true;
            else btnModule.IsChecked = true;

            conn = connector;
        }

        PPTConnector conn;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (btnClass.IsChecked.Value)
            { conn.AddClass(TBName.Text); }
            else
            { conn.AddModule(TBName.Text); }
            
            this.Close();
        }
    }
}
