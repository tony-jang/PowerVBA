using PowerVBA.Codes.TypeSystem;
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
        public List<Error> CodeErrors = new List<Error>();
        public ErrorList()
        {
            InitializeComponent();
        }
        public void SetError(List<Error> Errors)
        {
            CodeErrors = Errors;
            lvErrors.Items.Clear();
            CodeErrors.ForEach(err => { lvErrors.Items.Add(err.Message); });
        }
    }
}
