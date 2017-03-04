using System.Windows.Input;

namespace PowerVBA.Commands
{
    static class TabItemCommand
    {
        public static RoutedCommand DeleteCommand { get; }

        static TabItemCommand()
        {
            DeleteCommand = new RoutedCommand("Delete", typeof(TabItemCommand));
        }
    }
}
