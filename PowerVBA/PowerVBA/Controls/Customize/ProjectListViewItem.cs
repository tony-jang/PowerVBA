﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    class ProjectTemplateButton : Button
    {

        public ProjectTemplateButton()
        {
            this.Style = FindResource("ProjectListViewItemStyle") as Style;
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            this.ToolTip = new ToolTip() { Content = Description, StaysOpen = true };
        }

        public static DependencyProperty SourceProperty = DependencyHelper.Register();
        public static DependencyProperty DescriptionProperty = DependencyHelper.Register();


        public ImageSource Source
        {
            get { return (ImageSource)GetValue(SourceProperty); }
            set { SetValue(SourceProperty, value); }
        }

        public string Description
        {
            get { return (string)GetValue(DescriptionProperty); }
            set
            {
                SetValue(DescriptionProperty, value);
                this.ToolTip = new ToolTip() { Content = Description, StaysOpen = true };
            }
        }
    }
}
