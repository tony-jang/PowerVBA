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
    /// PreDeclareFuncManager.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PreDeclareFuncManager : UserControl
    {

        public PreDeclareFuncManager()
        {
            InitializeComponent();

            FileListView.SelectionChanged += FileListView_SelectionChanged;
            FunctionsListView.SelectionChanged += FunctionsListView_SelectionChanged;

            foreach (var func in PreDeclareFunctions.Functions)
            {
                bool FileExists = FileListView.Items.Cast<ListViewItem>()
                                                    .Where((itm) => itm.Content.ToString() == func.File).Count() != 0;

                if (!FileExists) FileListView.Items.Add(new ListViewItem() { Content = func.File });

                var FileLvItm = FileListView.Items.Cast<ListViewItem>()
                                                  .Where((itm) => itm.Content.ToString() == func.File).First();

                if (FileLvItm.Tag == null)
                {
                    var list = new List<PreDeclareFunction>();
                    FileLvItm.Tag = list;

                    list.Add(func);
                }
                else
                {
                    List<PreDeclareFunction> list = FileLvItm.Tag as List<PreDeclareFunction>;

                    list.Add(func);
                }
            }
        }

        private void FileListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<PreDeclareFunction> list = ((ListViewItem)FileListView.SelectedItem).Tag as List<PreDeclareFunction>;


            FunctionsListView.Items.Clear();

            foreach (var func in list)
            {
                string str = func.Identifier;
                if (func.ReturnData != "") str += $" ({func.ReturnData})";

                CheckBox cb = new CheckBox()
                {
                    Content = str,
                    IsChecked = func.IsUse,
                    Tag = func
                };
                cb.Checked += Cb_Checked;
                cb.Unchecked += Cb_Checked;

                FunctionsListView.Items.Add(cb);
                
            }
        }


        private void FunctionsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FunctionsListView.SelectedItem == null) return;
            PreDeclareFunction func = ((CheckBox)FunctionsListView.SelectedItem).Tag as PreDeclareFunction;

            RunFuncName.Text = func.File + "." + func.Identifier;
            TBDescription.Text = func.Description;
        }



        private void Cb_Checked(object sender, RoutedEventArgs e)
        {
            ListView lv = (ListView)(((CheckBox)sender).Parent);
            lv.SelectedItem = sender;
            ((PreDeclareFunction)((CheckBox)sender).Tag).IsUse = ((CheckBox)sender).IsChecked.Value;
        }

        public bool Saved { get; set; }
        public event BlankEventHandler SaveRequest;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveRequest?.Invoke();
            Saved = true;
        }
    }
}
