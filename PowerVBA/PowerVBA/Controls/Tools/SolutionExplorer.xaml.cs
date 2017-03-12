using PowerVBA.Core.Connector;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Controls.Customize;
using PowerVBA.Resources;
using PowerVBA.Global;
using PowerVBA.V2013.Wrap.WrapClass;
using PowerVBA.Core.Wrap.WrapBase;
using static PowerVBA.Wrap.WrappedClassManager;

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

        public void AddItem(VBComponentWrappingBase comp)
        {
            
            var t = comp.ToVBComponent2013().Type;

            ListBox AddLB = null;
            ImageSource img = null;
            
            if (t == VBA.vbext_ComponentType.vbext_ct_ClassModule) { AddLB = LBClass; img = ResourceImage.GetIconImage("ClassIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_StdModule) { AddLB = LBModule; img = ResourceImage.GetIconImage("ModuleIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_MSForm) { AddLB = LBForms; img = null; }

            AddLB?.Items.Add(new ImageListViewItem() { Content = comp.ToVBComponent2013().Name, Tag = comp, Source = img });
        }

        public void RemoveItem(VBComponentWrappingBase comp)
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


        public void Update(IPPTConnector pptConn)
        {
            IEnumerable<VBComponentWrappingBase> AddComp = new List<VBComponentWrappingBase>();
            IEnumerable<VBComponentWrappingBase> RemoveComp = new List<VBComponentWrappingBase>();
            
            var LocalItm = LBClass.Items.Cast<ImageListViewItem>()
                .Concat(LBForms.Items.Cast<ImageListViewItem>())
                .Concat(LBModule.Items.Cast<ImageListViewItem>())
                .Select(i => (VBComponentWrappingBase)i.Tag);


            // 버전별 분류
            IEnumerable<VBComponentWrappingBase> PPTItm = null;

            if (pptConn.GetType() == typeof(PowerVBA.V2013.Connector.PPTConnector2013))
            {
                var Conn2013 = (V2013.Connector.PPTConnector2013)pptConn;
                PPTItm = Conn2013.VBProject.VBComponents.Cast<VBA.VBComponent>().Select((i) => new VBComponentWrapping(i));
            }

            

            AddComp = PPTItm.Where((i) => !LocalItm.Contains(i)).Copy();
            RemoveComp = LocalItm.Where(i => !PPTItm.Contains(i)).Copy();

            foreach(var itm in AddComp) AddItem(itm);
            foreach (var itm in RemoveComp) RemoveItem(itm);
        }
    }
}
