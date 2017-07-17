using PowerVBA.Codes;
using PowerVBA.Codes.Parsing;
using PowerVBA.Controls.Customize;
using PowerVBA.Resources;
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
    /// Outline.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Outline : UserControl
    {
        public event EventHandler<LinePointEventArgs> LinePointMove;
        public Outline()
        {
            InitializeComponent();

            outlineItems.Items.Clear();
        }

        public void OnLinePointMove(LinePoint linePoint, string fileName)
        {
            LinePointMove?.Invoke(this, new LinePointEventArgs(linePoint, fileName));
        }

        public ImageTreeViewItem GetFileItem(string fileName)
        {
            ImageTreeViewItem fileItm;
            if (!outlineItems.Items.Cast<ImageTreeViewItem>().Select(i => i.Header.ToString()).Contains(fileName))
            {
                fileItm = new ImageTreeViewItem()
                {
                    Header = fileName,
                    Source = ResourceImage.GetIconImage("ModuleIcon"),
                };

                outlineItems.Items.Add(fileItm);
            }
            else
            {
                fileItm = outlineItems.Items.Cast<ImageTreeViewItem>().Where(i => i.Header.ToString() == fileName).FirstOrDefault();
            }

            return fileItm;
        }

        public ImageTreeViewItem GetItem(string itemName, string imageName, object tag = null)
        {
            var itm = new ImageTreeViewItem()
            {
                Header = itemName,
                Source = ResourceImage.GetIconImage(imageName),
                Tag = tag
            };
            itm.MouseDoubleClick += Itm_MouseDoubleClick;

            return itm;
        }

        public void AddVariable(Variable variable, string fileName)
        {
            var fileItm = GetFileItem(fileName);
            
            fileItm.Items.Add(GetItem(variable.Name, "VariableIcon", variable));
        }
        public void AddFunction(Function function, string fileName)
        {
            var fileItm = GetFileItem(fileName);

            fileItm.Items.Add(GetItem(function.Name, "MethodIcon", function));
        }
        public void AddSub(Sub sub, string fileName)
        {
            var fileItm = GetFileItem(fileName);
            fileItm.Items.Add(GetItem(sub.Name, "MethodSubIcon", sub));
        }
        public void AddEnum(EnumItem enumitem, string fileName)
        {
            var fileItm = GetFileItem(fileName);
            fileItm.Items.Add(GetItem(enumitem.Name, "EnumIcon", enumitem));
        }

        private void Itm_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (((Control)sender).Tag is IMember member)
            {
                OnLinePointMove(member.LinePoint, member.FileName);
            }
        }

        public void RemoveVariable(string VariableName)
        {
            var itm = outlineItems.Items
                         .Cast<ImageTreeViewItem>()
                         .Where(i => i.Header.ToString() == VariableName.ToString());

            outlineItems.Items.Remove(itm);
        }

        public void RemoveFile(params string[] fileNames)
        {
            foreach (var fileName in fileNames)
            {
                outlineItems.Items
                    .Cast<ImageTreeViewItem>()
                    .Where(i => i.Header.ToString() == fileName)
                    .ToList()
                    .ForEach(i => outlineItems.Items.Remove(i));
            }

        }

        public void ClearFile(params string[] fileNames)
        {
            foreach (var fileName in fileNames)
            {
                outlineItems.Items
                    .Cast<ImageTreeViewItem>()
                    .Where(i => i.Header.ToString() == fileName)
                    .ToList()
                    .ForEach(i => i.Items.Clear());
            }

        }

        public void ClearAll()
        {
            outlineItems.Items.Clear();
        }

        private void BtnChangeFilter_Click(object sender, MouseButtonEventArgs e)
        {

        }
    }
}
