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
using f=System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Vbe.Interop;

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// ReferenceManager.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ReferenceManager : UserControl
    {
        PPTConnectorBase Connector;

        public ReferenceManager(PPTConnectorBase Connector)
        {
            this.Connector = Connector;
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new f.OpenFileDialog();

            if (ofd.ShowDialog() == f.DialogResult.OK)
            {
                Connector.AddReference(ofd.FileName);
            }
        }

        public void AddItem(Reference reference)
        {
            
        }
    }
}
