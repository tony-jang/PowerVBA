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
using System.Runtime.InteropServices;
using PowerVBA.Core.Wrap.WrapBase;

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
            InitializeComponent();
            lvReferences.Items.Clear();

            this.Connector = Connector;
            foreach(var itm in Connector.References)
            {
                AddItem(itm);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new f.OpenFileDialog();

            if (ofd.ShowDialog() == f.DialogResult.OK)
            {
                if (Connector.AddReference(ofd.FileName))
                {
                    // TODO : AddReference 함수 변경 + 추가 구현
                }
            }
        }

        
        public void AddItem(ReferenceWrappingBase reference)
        {
            var cb = new CheckBox()
            {
                Content = reference.FullPath
            };

            cb.Checked += Cb_Checked;
            cb.Unchecked += Cb_Checked;
            lvReferences.Items.Add(cb);
        }

        private void Cb_Checked(object sender, RoutedEventArgs e)
        {
            
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
