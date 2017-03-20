using PowerVBA.Core.AvalonEdit.Errors;
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
    /// ErrorList.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ErrorList : UserControl
    {
        public List<CodeError> CodeErrors = new List<CodeError>();
        public ErrorList()
        {
            InitializeComponent();
            lvErrors.ItemsSource = CodeErrors;
            CodeErrors.Add(new CodeError("IDE1", "해당 문법은 현재 버전에서 지원하지 않습니다. 지원 문법을 확인하세요.", "Module1", 1));
        }
    }
}
