﻿using System;
using System.Collections.Generic;
using System.IO;
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
    /// PresentationAnalyzer.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProjectAnalyzer : UserControl
    {
        public ProjectAnalyzer()
        {
            InitializeComponent();
        }

        public event BlankEventHandler OpenShapeExplorer;

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

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenShapeExplorer?.Invoke();
        }
        
        public void SlideSync(V2013.Connector.PPTConnector2013 Connector)
        {
            RunSelSlideShapeCount.Text = Connector.Presentation
                                    .Slides[Connector.Application.ActiveWindow.Selection.SlideRange.SlideIndex]
                                    .Shapes.Count.ToString();
            RunSelShape.Text = Connector.Application.ActiveWindow.Selection.ShapeRange.Name;
        }

        public void CodeSync(int lines, int Components, int Currline)
        {
            RunAllLine.Text = lines.ToString();
            RunVBAClasses.Text = Components.ToString();

            RunCurrLine.Text = Currline.ToString();
        }
    }
}