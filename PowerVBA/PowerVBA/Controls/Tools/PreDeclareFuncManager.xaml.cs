using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Resources.Functions;
using PowerVBA.Wrap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
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

            Functions = FunctionReader.GetFunctions();


            foreach (var func in Functions)
            {
                bool FileExists = FileListView.Items.Cast<ListViewItem>()
                                                    .Where((itm) => itm.Name == func.Name.Folder).Count() != 0;

                if (!FileExists) FileListView.Items.Add(new ListViewItem() { Name = func.Name.Folder });

                var FileLvItm = FileListView.Items.Cast<ListViewItem>()
                                                  .Where((itm) => itm.Name == func.Name.Folder).First();



                // 파일이 새로 만들어졌을 경우
                if (FileLvItm.Tag == null)
                {
                    var list = new List<Function>();

                    var selList = list.Where(i => i.IsUse);

                    FileLvItm.Tag = list;

                    list.Add(func);

                    FileLvItm.Content = $"{FileLvItm.Name} ({selList.Count()}/{list.Count}개)";

                }
                // 파일이 이미 있을 경우
                else
                {
                    List<Function> list = FileLvItm.Tag as List<Function>;

                    var selList = list.Where(i => i.IsUse);

                    list.Add(func);

                    FileLvItm.Content = $"{FileLvItm.Name} ({selList.Count()}/{list.Count}개)";
                }
            }
        }

        List<Function> Functions;

        private void FileListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<Function> list = ((ListViewItem)FileListView.SelectedItem).Tag as List<Function>;


            FunctionsListView.Items.Clear();

            foreach (var func in list)
            {
                string str = func.Name.FileName;
                if (func.ReturnType != "") str += $" ({func.ReturnType})";

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
            Function func = ((CheckBox)FunctionsListView.SelectedItem).Tag as Function;

            RunFuncName.Text = func.Name;
            TBDescription.Text = func.Description;
        }



        private void Cb_Checked(object sender, RoutedEventArgs e)
        {
            ListView lv = (ListView)(((CheckBox)sender).Parent);
            lv.SelectedItem = sender;
            ((Function)((CheckBox)sender).Tag).IsUse = ((CheckBox)sender).IsChecked.Value;

            var FileLvItm = (ListViewItem)FileListView.SelectedItem;
            var list = FileLvItm.Tag as List<Function>;
            var selList = list.Where(i => i.IsUse);


            FileLvItm.Content = $"{FileLvItm.Name} ({selList.Count()}/{list.Count}개)";
        }
        
        public bool Saved { get; set; }
        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {

            List<Function> list = new List<Function>();
            FileListView.Items.Cast<ListViewItem>()
                              .ToList()
                              .Select((i) => (List<Function>)i.Tag)
                              .ToList()
                              .ForEach((lst) => list.AddRange(lst));
            list = list.Where((i) => i.IsUse).ToList();


            // 종속성 함수
            List<Function> DependencyFunction = new List<Function>();

            DependencyFunction = list.Where((i) => !string.IsNullOrEmpty(i.DependencyMessage)).ToList();

            if (DependencyFunction.Count >= 1)
            {
                foreach (var itm in DependencyFunction)
                {
                    MessageBox.Show($"종속성 정보를 발견했습니다.\r\n\r\n{itm.DependencyMessage}");
                }
            }


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
                if (lastFile != itm.Name.Folder)
                {
                    if (lastFile != "") sb.AppendLine($"'#endregion");
                    sb.AppendLine();
                    sb.AppendLine($"'#region \"Part Of {itm.Name.Folder}\"");

                    lastFile = itm.Name.Folder;
                }
                sb.AppendLine(itm.Code);
                
                sb.AppendLine();

                infoSb.AppendLine("'" + itm.Name);
            }

            sb.AppendLine($"'#endregion");

            infoSb.AppendLine("'[End]");

            sb.AppendLine();

            module.ToVBComponent2013().CodeModule.DeleteLines(1, module.ToVBComponent2013().CodeModule.CountOfLines);
            module.ToVBComponent2013().CodeModule.AddFromString(infoSb.ToString() + Environment.NewLine + sb.ToString());
        }

        public void SyncItem()
        {
            foreach (ListViewItem itm in FileListView.Items)
            {
                List<Function> list = itm.Tag as List<Function>;
                var selList = list.Where(i => i.IsUse);

                itm.Content = $"{itm.Name} ({selList.Count()}/{list.Count}개)";
            }

            foreach (CheckBox cb in FunctionsListView.Items)
            {
                Function func = cb.Tag as Function;
                cb.IsChecked = func.IsUse;
            }
        }


        // 전체 선택
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            Functions.ForEach(i => i.IsUse = true);
            SyncItem();
        }

        // 전체 선택 해제
        private void btnDeSelect_Click(object sender, RoutedEventArgs e)
        {
            Functions.ForEach(i => i.IsUse = false);
            SyncItem();
        }

        // 함수 전체 선택
        private void btnFuncSelect_Click(object sender, RoutedEventArgs e)
        {
            List<Function> list = ((ListViewItem)FileListView.SelectedItem).Tag as List<Function>;

            list.ForEach(i => i.IsUse = true);
            SyncItem();
        }

        // 함수 전체 해제
        private void btnFuncDeSelect_Click(object sender, RoutedEventArgs e)
        {
            List<Function> list = ((ListViewItem)FileListView.SelectedItem).Tag as List<Function>;

            list.ForEach(i => i.IsUse = false);
            SyncItem();
        }


        public int compare(Function f1, Function f2)
        {
            return f1.Name.ToString().CompareTo(f2.Name);
        }

        public PPTConnectorBase Connector { get; set; }

    }
}
