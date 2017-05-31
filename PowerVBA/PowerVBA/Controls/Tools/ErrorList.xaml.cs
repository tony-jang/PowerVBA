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
        public List<Error> codeErrors = new List<Error>();
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

                LineMove?.Invoke((Error)listview.Tag);
            }
        }

        public void Reset()
        {
            lvErrors.Items.Clear();
            codeErrors = new List<Error>();
        }

        public void SetError(List<Error> Errors)
        {
            codeErrors = Errors;
            SetErrorCtrl();
        }
        public void SetErrorCtrl()
        {
            lvErrors.Items.Clear();

            int errCount = codeErrors.Where(i => i.ErrorType == ErrorType.Error).Count(),
                warnCount = codeErrors.Where(i => i.ErrorType == ErrorType.Warning).Count();


            int errVisCount = 0, warnVisCount = 0;

            codeErrors.Where((i) => (i.Message.ToLower().Contains(tbSearchError.Text.ToLower()) || tbSearchError.Text == string.Empty)).ToList().ForEach(err => {
                lvErrors.Items.Add(
                    new ListViewItem()
                    {
                        Content = $"[{err.FileName}] {err.Message}| ({err.Line}번째 줄) <{err.ErrorCode.ToString()}>",
                        Tag = err
                    });
                if (err.ErrorType == ErrorType.Error) errVisCount++;
                else if (err.ErrorType == ErrorType.Warning) warnVisCount++;

            });

            

            if (lvErrors.Items.Count != codeErrors.Count)
            {
                runWarnCount.Text = $"({warnVisCount}/{warnCount})";
                runErrorCount.Text = $"({errVisCount}/{errCount})";
            }
            else
            {
                runErrorCount.Text = errCount.ToString();
                runWarnCount.Text = warnCount.ToString();
            }
        }
        public delegate void MoveRequestEventHandler(Error err);
        public event MoveRequestEventHandler LineMove;

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetErrorCtrl();
        }
    }
}
