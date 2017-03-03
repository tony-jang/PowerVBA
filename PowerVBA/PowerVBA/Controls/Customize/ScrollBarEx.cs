using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    class ScrollBarEx : ScrollBar
    {
        public static DependencyProperty IsMaximumProperty = DependencyHelper.Register();

        public bool IsMaximum
        {
            get
            {
                Console.WriteLine((Maximum == Value).ToString());
                return Maximum == Value; //(bool)GetValue(IsMaximumProperty);
            }
            set { SetValue(IsMaximumProperty, value); }
        }
    }
}
