using PowerVBA.Core.Connector;
using PowerVBA.Wrap;
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
    /// ProjectProperty.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProjectProperty : UserControl
    {
        public bool IsChanged = false;

        public ProjectProperty(PPTConnectorBase conn)
        {
            InitializeComponent();

            connector = conn;

            tbProjDesc.Text = connector.ProjectDescription;
            tbProjName.Text = connector.ProjectName;
        }

        PPTConnectorBase connector;

        private void FileTabControl_Changed(object sender, SelectionChangedEventArgs e)
        {
            
            //connector.ToConnector2010().Presentation.Password
        }

        private void tbProjDesc_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                connector.ProjectDescription = tbProjDesc.Text;
            }
            catch (Exception)
            {
                MessageBox.Show("프로젝트 설명문을 변경 하던 중 오류가 발생했습니다.");
            }
        }

        private void tbProjName_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                connector.ProjectName = tbProjName.Text;
            }
            catch (Exception)
            {
                MessageBox.Show("프로젝트 이름을 변경 하던 중 오류가 발생했습니다.");
            }
        }
    }
}
