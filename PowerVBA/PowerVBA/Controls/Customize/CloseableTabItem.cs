using PowerVBA.Commands;
using PowerVBA.Core.Connector;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPFExtension;

namespace PowerVBA.Controls.Customize
{
    public class CloseableTabItem : TabItem
    {

        public event HandledEventHandler SaveCloseRequest;

        public CloseableTabItem()
        {
            CommandBindings.Add(new CommandBinding(TabItemCommand.DeleteCommand, OnDelete, OnDeleteExecute));
        }

        public static DependencyProperty SavedProperty = DependencyHelper.Register(new PropertyMetadata(true));

        /// <summary>
        /// 저장되었는지에 대한 여부를 가져옵니다. 
        /// true는 저장됨 false는 저장 되지 않음을 나타냅니다.
        /// </summary>
        public bool Saved
        {
            get { return (bool)GetValue(SavedProperty); }
            set { SetValue(SavedProperty, value); }
        }

        private void OnDelete(object sender, ExecutedRoutedEventArgs e)
        {
            if (this.Parent.GetType() == typeof(TabControl))
            {
                if (!Saved)
                {
                    var result = MessageBox.Show("저장되지 않은 내용이 있습니다. 저장하고 닫으시겠습니까?", "저장되지 않은 파일", MessageBoxButton.YesNoCancel);

                    if (result == MessageBoxResult.Yes)
                    {
                        HandledEventArgs ev = new HandledEventArgs();
                        
                        SaveCloseRequest?.Invoke(this, ev);
                        
                        if (ev.Handled)
                            return;
                    }
                    else if (result == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                }
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
