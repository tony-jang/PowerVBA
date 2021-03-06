﻿using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using System.Threading;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Controls;

namespace PowerVBA.Controls.Tools
{
    /// <summary>
    /// SlideExplorer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ShapeExplorer : UserControl
    {
        PPTConnectorBase connector;

        public ShapeExplorer(PPTConnectorBase connector)
        {
            InitializeComponent();

            SetToUnknown();
            lvShape.SelectionChanged += LvShape_SelectionChanged;
            this.connector = connector;

            RefreshSlide();
        }

        public void SetToUnknown()
        {
            tbName.Text = "선택되지 않음";
            runType.Text = "알 수 없음";

            runWidth.Text = "NaN";
            runHeight.Text = "NaN";
            runTop.Text = "NaN";
            runLeft.Text = "NaN";
            runFill.Text = "알 수 없음";
            runTxtFill.Text = "알 수 없음";
        }
        private void RefreshSlide()
        {
            lvShape.Items.Clear();
            SetToUnknown();
            cbSlideList.Items.Clear();

            
            runSlide.Text = "0";
            
            for(int i = 0; i< connector.SlideCount; i++)
            {
                cbSlideList.Items.Add(new ComboBoxItem() { Content = (i + 1) + " 슬라이드" });
            }
        }
        private void RefreshShape(int slide)
        {
            lvShape.Items.Clear();

            Thread thr = new Thread(() =>
            {

                Dispatcher.Invoke(() =>
                {
                    runSlide.Text = slide.ToString();

                    tbMsg.Text = "파워포인트에서 도형 정보를 가져오고 있습니다.";
                    loadBorder.Visibility = Visibility.Visible;

                    tbLoad.Text = "대기중..";
                });
                var itm = connector.Shapes(slide);

                int counter = 0;

                foreach (ShapeWrappingBase shape in itm)
                {
                    Dispatcher.Invoke(() =>
                    {
                        tbLoad.Text = $"({++counter}/{itm.Count})";
                        lvShape.Items.Add(new ListViewItem() { Content = shape.Name, Tag = shape });
                    }, DispatcherPriority.Background);
                }

                Dispatcher.Invoke(() => { loadBorder.Visibility = Visibility.Hidden; });
            });

            thr.Start();
        }

        private void LvShape_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lvShape.SelectedIndex == -1)
                return;
            var tag = ((Control)lvShape.SelectedItem).Tag;

            if (tag == null)
                return;

            if (tag is ShapeWrappingBase shape)
            {
                tbName.Text = shape.Name;
                runType.Text = shape.ShapeType;
                runWidth.Text = shape.Width.ToString();
                runHeight.Text = shape.Height.ToString();
                runLeft.Text = shape.Left.ToString();
                runTop.Text = shape.Top.ToString();
                runFill.Text = $"R : {shape.RGB.R} | G : {shape.RGB.G} | B : {shape.RGB.B}";
                runTxtFill.Text = $"R : {shape.ForeRGB.R} | G : {shape.ForeRGB.G} | B : {shape.ForeRGB.B}";
            }
        }

        private void BtnRefresh_ButtonClick(object sender)
        {
            RefreshSlide();
        }

        private void BtnDelShape_ButtonClick(object sender)
        {
            if (lvShape.SelectedIndex == -1)
                return;

            var tag = ((Control)lvShape.SelectedItem).Tag;

            if (tag == null)
                return;

            if (tag is ShapeWrappingBase shape)
            {
                if (MessageBox.Show($"{shape.Name} 도형을 삭제합니다. 계속하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    shape.Delete(out bool success);
                    if (!success)
                    {
                        MessageBox.Show("삭제에 실패했습니다.");
                        return;
                    }
                    lvShape.Items.RemoveAt(lvShape.SelectedIndex);
                }
            }
        }

        private void CbSlideList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbSlideList.SelectedItem == null)
                return;

            if (cbSlideList.SelectedItem is ComboBoxItem cbItem)
            {
                int slide = int.Parse(cbItem.Content.ToString().Split(' ')[0]);

                RefreshShape(slide);
            }
        }
    }
}
