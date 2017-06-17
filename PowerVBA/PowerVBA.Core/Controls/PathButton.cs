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

        public static DependencyProperty FillProperty = DependencyHelper.Register();

        public static DependencyProperty UseMouseOverFillProperty = DependencyHelper.Register(new PropertyMetadata(false));
        public static DependencyProperty MouseOverFillProperty = DependencyHelper.Register();

        public static DependencyProperty UseMouseOverBackgroundProperty = DependencyHelper.Register(new PropertyMetadata(false));
        public static DependencyProperty MouseOverBackgroundProperty = DependencyHelper.Register();

        public Geometry Data
        {
            get => (Geometry)GetValue(DataProperty);
            set => SetValue(DataProperty, value);
        }

        public Brush Fill
        {
            get => (Brush)GetValue(FillProperty);
            set => SetValue(FillProperty, value);
        }

        public double ContentWidth
        {
            get => (double)GetValue(ContentWidthProperty);
            set => SetValue(ContentWidthProperty, value);
        }
        public double ContentHeight
        {
            get => (double)GetValue(ContentHeightProperty);
            set => SetValue(ContentHeightProperty, value);
        }
        public HorizontalAlignment ContentHorizontalAlignment
        {
            get => (HorizontalAlignment)GetValue(ContentHorizontalAlignmentProperty);
            set => SetValue(ContentHorizontalAlignmentProperty, value);
        }

        public VerticalAlignment ContentVerticalAlignment
        {
            get => (VerticalAlignment)GetValue(ContentVerticalAlignmentProperty);
            set => SetValue(ContentVerticalAlignmentProperty, value);
        }

        public Thickness ContentMargin
        {
            get => (Thickness)GetValue(ContentMarginProperty);
            set => SetValue(ContentMarginProperty, value);
        }

        public bool UseMouseOverFill
        {
            get => (bool)GetValue(UseMouseOverFillProperty);
            set => SetValue(UseMouseOverFillProperty, value);
        }

        public Brush MouseOverFill
        {
            get => (Brush)GetValue(MouseOverFillProperty);
            set => SetValue(MouseOverFillProperty, value);
        }

        public bool UseMouseOverBackground
        {
            get => (bool)GetValue(UseMouseOverBackgroundProperty);
            set => SetValue(UseMouseOverBackgroundProperty, value);
        }

        public Brush MouseOverBackground
        {
            get => (Brush)GetValue(MouseOverBackgroundProperty);
            set => SetValue(MouseOverBackgroundProperty, value);
        }
    }
}
