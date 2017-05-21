using System;
using System.Windows.Input;
using PowerVBA.Core.Connector;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Controls.Customize;
using PowerVBA.Resources;
using PowerVBA.Global;
using PowerVBA.Core.Wrap.WrapBase;
using static PowerVBA.Wrap.WrappedClassManager;
using static PowerVBA.Global.Globals;
using System.Windows;

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// SolutionExplorer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class SolutionExplorer : UserControl
    {
        
        public SolutionExplorer()
        {
            InitializeComponent();

            ListBoxes.Add(LBClass);
            ListBoxes.Add(LBModule);
            ListBoxes.Add(LBForms);
            ListBoxes.Add(LBSlideDoc);


            foreach (ListBox lb in ListBoxes)
            {
                lb.MouseDoubleClick += Item_DoubleClick;
            }


            // ContextMenu 초기화


            MenuItem itm1 = new MenuItem();
            MenuItem itm2 = new MenuItem();
            MenuItem itm3 = new MenuItem();

            itm1.Header = "열기";
            itm1.Click += Itm1_Click;
            itm2.Header = "복사";
            itm2.Click += Itm2_Click;
            itm3.Header = "삭제";
            itm3.Click += Itm3_Click;

            itmMenu.Items.Add(itm1);
            itmMenu.Items.Add(itm2);
            itmMenu.Items.Add(itm3);
        }

        public List<ListBox> ListBoxes = new List<ListBox>();

        ContextMenu itmMenu = new ContextMenu();


        private void Itm1_Click(object sender, RoutedEventArgs e)
        {
            var itm = (ImageListViewItem)GetSelectedItem();
            VBComponentWrappingBase comp = (VBComponentWrappingBase)itm.Tag;
            
            Open?.Invoke(this, comp);
        }

        private void Itm2_Click(object sender, RoutedEventArgs e)
        {
            var itm = (ImageListViewItem)GetSelectedItem();
            VBComponentWrappingBase comp = (VBComponentWrappingBase)itm.Tag;

            Copy?.Invoke(this, comp);
        }
        
        private void Itm3_Click(object sender, RoutedEventArgs e)
        {
            var itm = (ImageListViewItem)GetSelectedItem();
            VBComponentWrappingBase comp = (VBComponentWrappingBase)itm.Tag;
            if (MessageBox.Show($"'{itm.Content}'가 영구적으로 삭제됩니다.", "삭제 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Delete?.Invoke(this, comp);
            }
        }

        public delegate void ComponentDelegate(object sender, VBComponentWrappingBase data);

        public event ComponentDelegate Open;
        public event ComponentDelegate Copy;
        public event ComponentDelegate Delete;

        public event BlankDelegate OpenProperty;
        public event BlankDelegate OpenObjectBrowser;
        public event BlankDelegate OpenShapeExplorer;

        private void Item_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            object source = (((FrameworkElement)e.OriginalSource).TemplatedParent);
            if (source.GetType() == typeof(ContentPresenter)) source = ((ContentPresenter)source).TemplatedParent;

            ImageListViewItem itm = (ImageListViewItem)source;
            VBComponentWrappingBase comp = (VBComponentWrappingBase)itm.Tag;


            Open?.Invoke(this, comp);
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

        public ListBoxItem GetSelectedItem()
        {
            foreach (ListBox lb in ListBoxes)
            {
                foreach(ListBoxItem lbItm in lb.Items)
                {
                    if (lbItm.IsSelected) return lbItm;
                }
            }
            return null;
        }

        public void AddItem(VBComponentWrappingBase comp)
        {
            
            var t = comp.ToVBComponent2013().Type;

            ListBox AddLB = null;
            ImageSource img = null;
            
            if (t == VBA.vbext_ComponentType.vbext_ct_ClassModule) { AddLB = LBClass; img = ResourceImage.GetIconImage("ClassIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_StdModule) { AddLB = LBModule; img = ResourceImage.GetIconImage("ModuleIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_MSForm) { AddLB = LBForms; img = ResourceImage.GetIconImage("FormIcon"); }
            else if (t == VBA.vbext_ComponentType.vbext_ct_Document) { AddLB = LBSlideDoc; img = ResourceImage.GetIconImage("ClassIcon"); }

            var item = new ImageListViewItem() { Content = $"{comp.ToVBComponent2013().Name}{GetExtensions(t)}", Tag = comp, Source = img, ContextMenu = itmMenu };
            item.KeyDown += Item_KeyDown;
            AddLB?.Items.Add(item);
        }

        private void Item_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                Itm3_Click(sender, e);
            }
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

        public void UpdateSlide(IPPTConnector pptConn)
        {
            btnOpenShapeExplorer.Text = $"도형 탐색기 ({pptConn.Slide} 슬라이드)";
        }

        public void Update(IPPTConnector pptConn)
        {
            IEnumerable<VBComponentWrappingBase> AddComp = new List<VBComponentWrappingBase>();
            IEnumerable<VBComponentWrappingBase> RemoveComp = new List<VBComponentWrappingBase>();
            
            var LocalItm = LBClass.Items.Cast<ImageListViewItem>()
                .Concat(LBForms.Items.Cast<ImageListViewItem>())
                .Concat(LBModule.Items.Cast<ImageListViewItem>())
                .Concat(LBSlideDoc.Items.Cast<ImageListViewItem>())
                .Select(i => (VBComponentWrappingBase)i.Tag);


            // 버전별 분류
            IEnumerable<VBComponentWrappingBase> PPTItm = null;

            if (pptConn.GetType() == typeof(V2013.Connector.PPTConnector2013))
            {
                var Conn2013 = (V2013.Connector.PPTConnector2013)pptConn;
                PPTItm = Conn2013.VBProject.VBComponents.Cast<VBA.VBComponent>().Select((i) => new V2013.WrapClass.VBComponentWrapping(i));
            }
            else if (pptConn.GetType() == typeof(V2010.Connector.PPTConnector2010))
            {
                var Conn2013 = (V2010.Connector.PPTConnector2010)pptConn;
                PPTItm = Conn2013.VBProject.VBComponents.Cast<VBA.VBComponent>().Select((i) => new V2010.WrapClass.VBComponentWrapping(i));
            }


            AddComp = PPTItm.Where((i) => !LocalItm.Contains(i)).Copy();
            RemoveComp = LocalItm.Where(i => !PPTItm.Contains(i)).Copy();

            foreach(var itm in AddComp) AddItem(itm);
            foreach (var itm in RemoveComp) RemoveItem(itm);
            
            ClassRun.Text = LBClass.Items.Count.ToString();
            ModuleRun.Text = LBModule.Items.Count.ToString();
            FormRun.Text = LBForms.Items.Count.ToString();
            SlideDocRun.Text = LBSlideDoc.Items.Count.ToString();
        }

        private void OpenProperty_Click(object sender, MouseButtonEventArgs e)
        {
            OpenProperty?.Invoke();
        }

        private void BtnOpenObjectBrowser_Click(object sender, MouseButtonEventArgs e)
        {
            OpenObjectBrowser?.Invoke();
        }

        private void OpenShapeExplorer_Click(object sender, MouseButtonEventArgs e)
        {
            OpenShapeExplorer?.Invoke();
        }
    }
}
