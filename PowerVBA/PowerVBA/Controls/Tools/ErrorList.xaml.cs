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
using static PowerVBA.Global.Globals;

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
            lvErrors.MouseDoubleClick += LvErrors_MouseDoubleClick;
        }

        private void LvErrors_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var itm = ((FrameworkElement)e.OriginalSource).TemplatedParent;
            if (itm.GetType() == typeof(ContentPresenter)) itm = ((FrameworkElement)itm).TemplatedParent;

            if (itm.GetType() == typeof(ListViewItem))
            {
                ListViewItem listview = (ListViewItem)itm;

                LineMoveRequest?.Invoke((Error)listview.Tag);
            }
        }

        public void SetError(List<Error> Errors)
        {
            CodeErrors = Errors;
            SetErrorCtrl();
        }
        public void SetErrorCtrl()
        {
            lvErrors.Items.Clear();

            CodeErrors.Where((i) => (i.Message.Contains(TBSearchError.Text) || TBSearchError.Text == string.Empty)).ToList().ForEach(err => {
                lvErrors.Items.Add(
                    new ListViewItem()
                    {
                        Content = err.Message + " | (" + err.Line + "번째 줄) <" + err.ErrorCode.ToString() + ">",
                        Tag = err
                    }
                    );
            });

            if (lvErrors.Items.Count != CodeErrors.Count)
                RunErrorCount.Text = $"({lvErrors.Items.Count}/{CodeErrors.Count})";
            else
                RunErrorCount.Text = CodeErrors.Count.ToString();
        }
        public delegate void MoveRequestEventHandler(Error err);
        public event MoveRequestEventHandler LineMoveRequest;

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetErrorCtrl();
        }

        
    }
}
