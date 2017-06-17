using System.Windows.Input;

namespace PowerVBA.Commands
{
    static class WindowSystemCommand
    {
        public static RoutedCommand CloseCommand { get; }
        public static RoutedCommand RestoreCommand { get; }
        public static RoutedCommand MinimizeCommand { get; }
        public static RoutedCommand HelpCommand { get; }

        static WindowSystemCommand()
        {
            CloseCommand = new RoutedCommand("Close", typeof(WindowSystemCommand));
            RestoreCommand = new RoutedCommand("Restore", typeof(WindowSystemCommand));
            MinimizeCommand = new RoutedCommand("Minimize", typeof(WindowSystemCommand));
            HelpCommand = new RoutedCommand("Help", typeof(WindowSystemCommand));
        }
    }
}
