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
using ppt = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Shapes;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Extension;

namespace PowerVBA.Windows
{
    /// <summary>
    /// ConnectWindows.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ConnectWindows : ChromeWindow
    {
        public ConnectWindows()
        {
            InitializeComponent();
            RefreshList();

            var refreshComm = new RoutedCommand();
            refreshComm.InputGestures.Add(new KeyGesture(Key.F5));

            CommandBindings.Add(new CommandBinding(refreshComm, refreshCommand));
        }

        private void refreshCommand(object sender, ExecutedRoutedEventArgs e)
        {
            RefreshList();
        }
        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            RefreshList();
        }

        public void RefreshList()
        {
            PptList.Items.Clear();

            var itm = new ppt.Application();

            foreach (ppt.Presentation presentation in itm.Presentations)
            {
                PptList.Items.Add(new ListViewItem() { Content = presentation.Name + (((Bool2)presentation.ReadOnly) ? " [읽기 전용]" : string.Empty),
                                                       Tag = presentation });
            }
        }

        PresentationWrappingBase returnPpt = null;
        bool IsHandled = false;
        public bool ShowDialog(out PresentationWrappingBase ppt)
        {
            base.ShowDialog();
            ppt = returnPpt;
            return IsHandled;
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            IsHandled = false;
            this.Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            if (PptList.SelectedIndex == -1)
            {
                MessageBox.Show("선택된 프레젠테이션이 없습니다.");
                return;
            }
            else
            {
                try
                {
                    ppt.Presentation ppt = ((ppt.Presentation)((ListViewItem)PptList.SelectedItem).Tag);
                    if (ppt.Name != "")
                    {
                        PPTVersion ver = VersionSelector.GetPPTVersion();

                        switch (ver)
                        {
                            case PPTVersion.PPT2010:
                                returnPpt = new V2010.WrapClass.PresentationWrapping(ppt);
                                break;
                            case PPTVersion.PPT2013:
                                returnPpt = new V2013.WrapClass.PresentationWrapping(ppt);
                                break;
                        }
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("선택된 프레젠테이션은 더 이상 존재하지 않습니다.");
                    return;
                }
            }

            IsHandled = true;
            this.Close();
        }
    }
}
