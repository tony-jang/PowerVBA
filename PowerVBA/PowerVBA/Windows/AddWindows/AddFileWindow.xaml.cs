using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
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
        public AddFileWindow(IPPTConnector connector, AddType AddType)
        {
            InitializeComponent();
            if (AddType == AddType.Class) btnClass.IsChecked = true;
            else if (AddType == AddType.Module) btnModule.IsChecked = true;
            else if(AddType == AddType.Form) btnForm.IsChecked = true;

            conn = connector;
        }

        IPPTConnector conn;

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            VBComponentWrappingBase wrap;
            bool Result = false;
            if (btnClass.IsChecked.Value)
            { Result = conn.AddClass(TBName.Text, out wrap); }
            else if (btnModule.IsChecked.Value)
            { Result = conn.AddModule(TBName.Text, out wrap); }
            else if (btnForm.IsChecked.Value)
            { Result = conn.AddForm(TBName.Text, out wrap); }

            if (!Result)
            { MessageBox.Show("파일 추가에 실패했습니다. 명명 규칙이 잘못 되었습니다.\r\n명명 규칙은 다음과 같습니다.\r\n - 이름의 시작은 문자만 올 수 있습니다.\r\n - 이외에는 빈칸을 제외한 _이나 문자 또는 숫자가 올 수 있습니다."); }
            else this.Close();
        }

        public enum AddType
        {
            /// <summary>
            /// Class를 추가합니다.
            /// </summary>
            Class,
            /// <summary>
            /// 모듈을 추가합니다.
            /// </summary>
            Module,
            /// <summary>
            /// 사용자 폼을 추가합니다.
            /// </summary>
            Form
        }
    }
}
