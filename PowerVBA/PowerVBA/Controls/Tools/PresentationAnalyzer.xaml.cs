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
    /// PresentationAnalyzer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PresentationAnalyzer : UserControl
    {
        public PresentationAnalyzer()
        {
            InitializeComponent();
        }

        private int _SlideCount;
        public int SlideCount
        {
            get { return _SlideCount; }
            set
            {
                _SlideCount = value;

                SlideRun.Text = _SlideCount.ToString();
            }
        }

        private int _ShapeCount;
        public int ShapeCount
        {
            get { return _ShapeCount; }
            set
            {
                _ShapeCount = value;

                ShapeRun.Text = _ShapeCount.ToString();
            }
        }

        //private int _SectionCount;
        //public int SectionCount
        //{
        //    get { return _SectionCount; }
        //    set
        //    {
        //        _SectionCount = value;

        //        SectionRun.Text = _SectionCount.ToString();
        //    }
        //}
    }
}
