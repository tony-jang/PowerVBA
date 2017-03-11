using PowerVBA.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerVBA.Controls.Customize
{
    class CloseableTabItem : TabItem
    {
        public CloseableTabItem()
        {
            CommandBindings.Add(new CommandBinding(TabItemCommand.DeleteCommand, OnDelete, OnDeleteExecute));
        }

        private void OnDelete(object sender, ExecutedRoutedEventArgs e)
        {
            if (this.Parent.GetType() == typeof(TabControl))
            {
                TabControl tc = (TabControl)this.Parent;
                tc.Items.Remove(this);
            }
        }

        private void OnDeleteExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }
    }
}
