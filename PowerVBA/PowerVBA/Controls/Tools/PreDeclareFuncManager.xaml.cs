using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
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
        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {

            List<PreDeclareFunction> list = new List<PreDeclareFunction>();
            FileListView.Items.Cast<ListViewItem>()
                              .ToList()
                              .Select((i) => (List<PreDeclareFunction>)i.Tag)
                              .ToList()
                              .ForEach((lst) => list.AddRange(lst));
            list = list.Where((i) => i.IsUse).ToList();

            VBComponentWrappingBase module;
            if (!Connector.ContainsModule("PowerVBA"))
            {
                Connector.AddModule("PowerVBA", out module);
            }
            else { module = Connector.GetModule("PowerVBA"); }
            
            if (module == null) MessageBox.Show("알 수 없는 오류가 발생했습니다.");

            StringBuilder sb = new StringBuilder();
            StringBuilder infoSb = new StringBuilder();
            list.Sort(compare);
            string lastFile = "";

            infoSb.AppendLine($"'          ┌───────────────────────────────┐");
            infoSb.AppendLine($"'          │                                                              │");
            infoSb.AppendLine($"'          │                 PowerVBA PreDeclare Functions                │");
            infoSb.AppendLine($"'          │                                                              │");
            infoSb.AppendLine($"'          │                                    [Auto-Generated Code]     │");
            infoSb.AppendLine($"'          └───────────────────────────────┘");
            infoSb.AppendLine($"' 이 부분은 PowerVBA에서 사용중인 프로시져들을 인식하기 위함입니다. 삭제하지 말아주세요.");

            infoSb.AppendLine("'[Start]");

            foreach (var itm in list)
            {
                if (lastFile != itm.File)
                {
                    if (lastFile != "") sb.AppendLine($"'#endregion");
                    sb.AppendLine();
                    sb.AppendLine($"'#region \"Part Of {itm.File}\"");

                    lastFile = itm.File;
                }
                sb.AppendLine(itm.Code);
                
                sb.AppendLine();

                infoSb.AppendLine("'" + itm.File + "." + itm.Identifier);
            }

            sb.AppendLine($"'#endregion");

            infoSb.AppendLine("'[End]");

            sb.AppendLine();

            module.ToVBComponent2013().CodeModule.DeleteLines(1, module.ToVBComponent2013().CodeModule.CountOfLines);
            module.ToVBComponent2013().CodeModule.AddFromString(infoSb.ToString() + Environment.NewLine + sb.ToString());
        }

        public int compare(PreDeclareFunction f1, PreDeclareFunction f2)
        {
            return f1.File.CompareTo(f2.File);
        }

        public PPTConnectorBase Connector { get; set; }

    }
}
