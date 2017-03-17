using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WPFExtension;

namespace PowerVBA.Core.Controls
{
    public class PathButton : Button
    {
        public static DependencyProperty DataProperty = DependencyHelper.Register();
        public static DependencyProperty ContentWidthProperty = DependencyHelper.Register();
        public static DependencyProperty ContentHeightProperty = DependencyHelper.Register();
        public static DependencyProperty ContentHorizontalAlignmentProperty = DependencyHelper.Register();
        public static DependencyProperty ContentVerticalAlignmentProperty = DependencyHelper.Register();
        public static DependencyProperty ContentMarginProperty = DependencyHelper.Register();
        public static DependencyProperty UseMouseOverColorProperty = DependencyHelper.Register();
        public static DependencyProperty MouseOverColorProperty = DependencyHelper.Register();

        public Geometry Data
        {
            get { return (Geometry)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        public double ContentWidth
        {
            get { return (double)GetValue(ContentWidthProperty); }
            set { SetValue(ContentWidthProperty, value); }
        }
        public double ContentHeight
        {
            get { return (double)GetValue(ContentHeightProperty); }
            set { SetValue(ContentHeightProperty, value); }
        }
        public HorizontalAlignment ContentHorizontalAlignment
        {
            get { return (HorizontalAlignment)GetValue(ContentHorizontalAlignmentProperty); }
            set { SetValue(ContentHorizontalAlignmentProperty, value); }
        }

        public VerticalAlignment ContentVerticalAlignment
        {
            get { return (VerticalAlignment)GetValue(ContentVerticalAlignmentProperty); }
            set { SetValue(ContentVerticalAlignmentProperty, value); }
        }

        public Thickness ContentMargin
        {
            get { return (Thickness)GetValue(ContentMarginProperty); }
            set { SetValue(ContentMarginProperty, value); }
        }

        public bool UseMouseOverColor
        {
            get { return (bool)GetValue(UseMouseOverColorProperty); }
            set { SetValue(UseMouseOverColorProperty, value); }
        }

        public Brush MouseOverColor
        {
            get { return (Brush)GetValue(MouseOverColorProperty); }
            set { SetValue(MouseOverColorProperty, value); }
        }
    }
}
