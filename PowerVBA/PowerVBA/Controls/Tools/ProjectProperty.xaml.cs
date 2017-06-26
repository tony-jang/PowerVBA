using PowerVBA.Controls.Customize;
using PowerVBA.Core.Connector;
using PowerVBA.Wrap;
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

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// ProjectProperty.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProjectProperty : UserControl
    {
        public bool IsChanged = false;

        public ProjectProperty(PPTConnectorBase conn, ref bool Error)
        {
            InitializeComponent();

            connector = conn;
            try
            {
                tbProjDesc.Text = connector.ProjectDescription;
                tbProjName.Text = connector.ProjectName;
            }
            catch (Exception)
            {
                Error = true;
                return;
            }
            
        }

        public void SavecloseRequest(object sender, HandledEventArgs e)
        {
            try
            {
                connector.ProjectName = tbProjName.Text;
                connector.ProjectDescription = tbProjDesc.Text;
            }
            catch (Exception)
            {
                MessageBox.Show("저장하던 중 오류가 발생했습니다. PowerPoint에서 사용하고 있는 것 같습니다.");
                e.Handled = true;
            }
        }

        PPTConnectorBase connector;

        private void FileTabControl_Changed(object sender, SelectionChangedEventArgs e)
        {
            
        }

        private void tbProjDesc_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.Parent is CloseableTabItem closeItm)
            {
                closeItm.Saved = false;
            }
        }

        private void tbProjName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.Parent is CloseableTabItem closeItm)
            {
                closeItm.Saved = false;
            }
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                connector.ProjectName = tbProjName.Text;
                connector.ProjectDescription = tbProjDesc.Text;
            }
            catch (Exception)
            {
                MessageBox.Show("프로젝트 정보을 변경 하던 중 오류가 발생했습니다.");
            }
        }
    }
}
