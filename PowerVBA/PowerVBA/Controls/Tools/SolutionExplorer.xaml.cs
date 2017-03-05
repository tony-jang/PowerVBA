using PowerVBA.Core.Connector;
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
using Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Core.Wrap.WrapClass;
using PowerVBA.Controls.Customize;
using PowerVBA.Resources;
using PowerVBA.Global;

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// SolutionExplorer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class SolutionExplorer : UserControl
    {
        public List<ListBox> ListBoxes = new List<ListBox>();
        public SolutionExplorer()
        {
            InitializeComponent();

            ListBoxes.Add(LBClass);
            ListBoxes.Add(LBModule);
            ListBoxes.Add(LBForms);
        }

        bool Handled = false;

        private void ListBoxes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Handled)
            {
                Handled = false;
                return;
            }

            foreach (ListBox lb in ListBoxes)
            {
                if (sender == lb) continue;

                if (lb.SelectedIndex != -1)
                {
                    Handled = true;
                    lb.SelectedIndex = -1;
                }
            }
        }

        public void AddItem(VBComponentWrapping comp)
        {
            var t = comp.Type;

            ListBox AddLB = null;
            ImageSource img = null;
            
            if (t == VBA.vbext_ComponentType.vbext_ct_ClassModule) { AddLB = LBClass; img = ResourceImage.GetIconImage("ClassIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_StdModule) { AddLB = LBModule; img = ResourceImage.GetIconImage("ModuleIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_MSForm) { AddLB = LBForms; img = null; }

            AddLB?.Items.Add(new ImageListViewItem() { Content = comp.Name, Tag = comp, Source = img });
        }

        public void RemoveItem(VBComponentWrapping comp)
        {
            
            foreach(ListBox lb in ListBoxes)
            {
                foreach(ImageListViewItem itm in lb.Items)
                {
                    if (itm.Tag == comp)
                    {
                        lb.Items.Remove(lb.Items.Cast<ImageListViewItem>().Where(i => i.Tag == comp).First());
                        return;            
                    }
                }
            }

        }


        public void Update(PPTConnector pptConn)
        {
            IEnumerable<VBComponentWrapping> AddComp = new List<VBComponentWrapping>();
            IEnumerable<VBComponentWrapping> RemoveComp = new List<VBComponentWrapping>();
            
            var LocalItm = LBClass.Items.Cast<ImageListViewItem>()
                .Concat(LBForms.Items.Cast<ImageListViewItem>())
                .Concat(LBModule.Items.Cast<ImageListViewItem>())
                .Select(i => (VBComponentWrapping)i.Tag);

            var PPTItm = pptConn.VBProject.VBComponents.Cast<VBA.VBComponent>().Select((i) => new VBComponentWrapping(i));

            AddComp = PPTItm.Where((i) => !LocalItm.Contains(i)).Copy();
            RemoveComp = LocalItm.Where(i => !PPTItm.Contains(i)).Copy();

            foreach(var itm in AddComp) AddItem(itm);
            foreach (var itm in RemoveComp) RemoveItem(itm);
        }
    }
}
